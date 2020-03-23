unit fuQImport3FormatsEditor;
{$I QImport3VerCtrl.Inc}
interface

uses
  {$IFDEF VCL16}
    Vcl.Forms,
    Data.Db,
    Vcl.StdCtrls,
    Vcl.Controls,
    Vcl.ComCtrls,
    System.Classes,
    Vcl.ExtCtrls,
    Vcl.ImgList,
    Vcl.ToolWin,
  {$ELSE}
    Forms,
    Db,
    StdCtrls,
    Controls,
    ComCtrls,
    Classes,
    ExtCtrls,
    ImgList,
    ToolWin,
  {$ENDIF}
  QImport3;

type
  TfmQImport3FormatsEditor = class(TForm)
    paButtons: TPanel;
    bOk: TButton;
    bCancel: TButton;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    lstFields: TListView;
    Bevel6: TBevel;
    ilFields: TImageList;
    pgFieldOptions: TPageControl;
    tsFieldTuning: TTabSheet;
    Bevel13: TBevel;
    laGeneratorValue: TLabel;
    laGeneratorStep: TLabel;
    laConstantValue: TLabel;
    laNullValue: TLabel;
    laDefaultValue: TLabel;
    laLeftQuote: TLabel;
    laRightQuote: TLabel;
    laQuoteAction: TLabel;
    laCharCase: TLabel;
    laCharSet: TLabel;
    edGeneratorValue: TEdit;
    edGeneratorStep: TEdit;
    edConstantValue: TEdit;
    edNullValue: TEdit;
    edDefaultValue: TEdit;
    edLeftQuote: TEdit;
    edRightQuote: TEdit;
    cbQuoteAction: TComboBox;
    cbCharCase: TComboBox;
    cbCharSet: TComboBox;
    tbReplacements: TToolBar;
    tbtAddReplacement: TToolButton;
    tbtEditReplacement: TToolButton;
    tbtDelReplacement: TToolButton;
    laReplacements: TLabel;
    lvReplacements: TListView;
    procedure FormShow(Sender: TObject);
    procedure edGeneratorValueChange(Sender: TObject);
    procedure edGeneratorStepChange(Sender: TObject);
    procedure edConstantValueChange(Sender: TObject);
    procedure edNullValueChange(Sender: TObject);
    procedure edDefaultValueChange(Sender: TObject);
    procedure cbCharCaseChange(Sender: TObject);
    procedure cbCharSetChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure lstFieldsEdited(Sender: TObject; Item: TListItem;
      var S: String);
    procedure lstFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure FormCreate(Sender: TObject);
    procedure lvReplacementsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure lvReplacementsDblClick(Sender: TObject);
    procedure tbtAddReplacementClick(Sender: TObject);
    procedure tbtEditReplacementClick(Sender: TObject);
    procedure tbtDelReplacementClick(Sender: TObject);
    procedure edLeftQuoteChange(Sender: TObject);
    procedure edRightQuoteChange(Sender: TObject);
    procedure cbQuoteActionChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    FComponent: TComponent;

    FIsNecessaryListChange: Boolean;
    FLoadingFormatItem: boolean;
    FFormatItem: TListItem;

    function GetGeneratorValue: integer;
    function GetGeneratorStep: integer;
    function GetConstantValue: string;
    function GetNullValue: string;
    function GetDefaultValue: string;
    function GetLeftQuote: string;
    function GetRightQuote: string;
    function GetQuoteAction: TQuoteAction;
    function GetCharCase: TQImportCharCase;
    function GetCharSet: TQImportCharSet;

    procedure SetGeneratorValue(const Value: integer);
    procedure SetGeneratorStep(const Value: integer);
    procedure SetConstantValue(const Value: string);
    procedure SetNullValue(const Value: string);
    procedure SetDefaultValue(const Value: string);
    procedure SetLeftQuote(const Value: string);
    procedure SetRightQuote(const Value: string);
    procedure SetQuoteAction(const Value: TQuoteAction);
    procedure SetCharCase(const Value: TQImportCharCase);
    procedure SetCharSet(const Value: TQImportCharSet);

    procedure FillFieldsList;
    procedure ApplyChanges;
    function DataByName(const AName: string): TQImportFieldFormat;
    procedure SetEnabledControls;
    function GetQImportFieldFormats: TQImportFieldFormats;
    function GetDataSet: TDataSet;
    procedure ShowFormatItem(Item: TListItem);
    procedure TuneButtons;
  public
    property IsNecessaryListChange: Boolean read FIsNecessaryListChange
      write FIsNecessaryListChange default True;
    property FieldFormats: TQImportFieldFormats read GetQImportFieldFormats;
    property DataSet: TDataSet read GetDataSet;
    property GeneratorValue: integer read GetGeneratorValue
      write SetGeneratorValue;
    property GeneratorStep: integer read GetGeneratorStep
      write SetGeneratorStep;
    property ConstantValue: string read GetConstantValue
      write SetConstantValue;
    property NullValue: string read GetNullValue
      write SetNullValue;
    property DefaultValue: string read GetDefaultValue
      write SetDefaultValue;
    property LeftQuote: string read GetLeftQuote
      write SetLeftQuote;
    property RightQuote: string read GetRightQuote
      write SetRightQuote;
    property QuoteAction: TQuoteAction read GetQuoteAction
      write SetQuoteAction;
    property CharCase: TQImportCharCase read GetCharCase
      write SetCharCase;
    property CharSet: TQImportCharSet read GetCharSet
      write SetCharSet;
  end;

function RunFormatsEditor(AComponent: TComponent): boolean;

implementation

uses
  {$IFDEF VCL16}
    System.TypInfo,
    System.SysUtils,
    Winapi.Windows,
  {$ELSE}
    TypInfo,
    SysUtils,
    Windows,
  {$ENDIF}
  QImport3StrIDs,
  fuQImport3ReplacementEdit;

{$R *.DFM}

function RunFormatsEditor(AComponent: TComponent): boolean;
begin
  with TfmQImport3FormatsEditor.Create(nil) do
  try
    FComponent := AComponent;
    FillFieldsList;
    IsNecessaryListChange := False;
    Result := ShowModal = mrOk;
    if Result then ApplyChanges;
  finally
    Free;
  end;
end;

{ TfmQImport3FormatsEditor }

function TfmQImport3FormatsEditor.GetGeneratorValue: integer;
begin
  Result := StrToIntDef(edGeneratorValue.Text, 0);
end;

function TfmQImport3FormatsEditor.GetGeneratorStep: integer;
begin
  Result := StrToIntDef(edGeneratorStep.Text, 0);
end;

function TfmQImport3FormatsEditor.GetConstantValue: string;
begin
  Result := edConstantValue.Text;
end;

function TfmQImport3FormatsEditor.GetNullValue: string;
begin
  Result := edNullValue.Text;
end;

function TfmQImport3FormatsEditor.GetDefaultValue: string;
begin
  Result := edDefaultValue.Text;
end;

function TfmQImport3FormatsEditor.GetLeftQuote: string;
begin
  Result := edLeftQuote.Text;
end;

function TfmQImport3FormatsEditor.GetRightQuote: string;
begin
  Result := edRightQuote.Text;
end;

function TfmQImport3FormatsEditor.GetQuoteAction: TQuoteAction;
begin
  case cbQuoteAction.ItemIndex of
    1: Result := qaAdd;
    2: Result := qaRemove;
  else
    Result := qaNone;
  end;
end;

function TfmQImport3FormatsEditor.GetCharCase: TQImportCharCase;
begin
  case cbCharCase.ItemIndex of
    1: Result := iccUpper;
    2: Result := iccLower;
    3: Result := iccUpperFirst;
    4: Result := iccUpperFirstWord;
  else
    Result := iccNone;
  end;
end;

function TfmQImport3FormatsEditor.GetCharSet: TQImportCharSet;
begin
  case cbCharSet.ItemIndex of
    1: Result := icsAnsi;
    2: Result := icsOem;
  else
    Result := icsNone;
  end;
end;

procedure TfmQImport3FormatsEditor.SetGeneratorValue(const Value: integer);
begin
  edGeneratorValue.Text := IntToStr(Value);
end;

procedure TfmQImport3FormatsEditor.SetGeneratorStep(const Value: integer);
begin
  edGeneratorStep.Text := IntToStr(Value);
end;

procedure TfmQImport3FormatsEditor.SetConstantValue(const Value: string);
begin
  edConstantValue.Text := Value;
end;

procedure TfmQImport3FormatsEditor.SetNullValue(const Value: string);
begin
  edNullValue.Text := Value;
end;

procedure TfmQImport3FormatsEditor.SetDefaultValue(const Value: string);
begin
  edDefaultValue.Text := Value;
end;

procedure TfmQImport3FormatsEditor.SetLeftQuote(const Value: string);
begin
  edLeftQuote.Text := Value;
end;

procedure TfmQImport3FormatsEditor.SetRightQuote(const Value: string);
begin
  edRightQuote.Text := Value;
end;

procedure TfmQImport3FormatsEditor.SetQuoteAction(const Value: TQuoteAction);
begin
  case Value of
    qaAdd: cbQuoteAction.ItemIndex := 1;
    qaRemove: cbQuoteAction.ItemIndex := 2;
  else
    cbQuoteAction.ItemIndex := 0;
  end;
end;

procedure TfmQImport3FormatsEditor.SetCharCase(const Value: TQImportCharCase);
begin
  case Value of
    iccUpper: cbCharCase.ItemIndex := 1;
    iccLower: cbCharCase.ItemIndex := 2;
    iccUpperFirst: cbCharCase.ItemIndex := 3;
    iccUpperFirstWord: cbCharCase.ItemIndex := 4;
  else
    cbCharCase.ItemIndex := 0;
  end;
end;

procedure TfmQImport3FormatsEditor.SetCharSet(const Value: TQImportCharSet);
begin
  case Value of
    icsAnsi: cbCharSet.ItemIndex := 1;
    icsOem: cbCharSet.ItemIndex := 2;
  else
    cbCharSet.ItemIndex := 0;
  end;
end;

procedure TfmQImport3FormatsEditor.FormShow(Sender: TObject);
begin
  FIsNecessaryListChange := True;
  Caption := FComponent.Name + '.DataFormats - Property editor';
  SetEnabledControls;
end;

procedure TfmQImport3FormatsEditor.edGeneratorValueChange(Sender: TObject);
begin
  if Assigned(lstFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFields.Selected.Data).GeneratorValue := GeneratorValue;
    SetEnabledControls;
  end;
end;

procedure TfmQImport3FormatsEditor.edGeneratorStepChange(Sender: TObject);
begin
  if Assigned(lstFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFields.Selected.Data).GeneratorStep := GeneratorStep;
    SetEnabledControls;
  end;
end;

procedure TfmQImport3FormatsEditor.edConstantValueChange(Sender: TObject);
begin
  if Assigned(lstFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFields.Selected.Data).ConstantValue := ConstantValue;
    SetEnabledControls;
  end;
end;

procedure TfmQImport3FormatsEditor.edNullValueChange(Sender: TObject);
begin
  if Assigned(lstFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFields.Selected.Data).NullValue := NullValue;
    SetEnabledControls;
  end;
end;

procedure TfmQImport3FormatsEditor.edDefaultValueChange(Sender: TObject);
begin
  if Assigned(lstFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFields.Selected.Data).DefaultValue := DefaultValue;
    SetEnabledControls;
  end;
end;

procedure TfmQImport3FormatsEditor.edLeftQuoteChange(Sender: TObject);
begin
  if Assigned(lstFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFields.Selected.Data).LeftQuote := LeftQuote;
    SetEnabledControls;
  end;
end;

procedure TfmQImport3FormatsEditor.edRightQuoteChange(Sender: TObject);
begin
  if Assigned(lstFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFields.Selected.Data).RightQuote := RightQuote;
    SetEnabledControls;
  end;
end;

procedure TfmQImport3FormatsEditor.cbQuoteActionChange(Sender: TObject);
begin
  if Assigned(lstFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFields.Selected.Data).QuoteAction := QuoteAction;
    SetEnabledControls;
  end;
end;

procedure TfmQImport3FormatsEditor.cbCharCaseChange(Sender: TObject);
begin
  if Assigned(lstFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFields.Selected.Data).CharCase := CharCase;
    SetEnabledControls;
  end;
end;

procedure TfmQImport3FormatsEditor.cbCharSetChange(Sender: TObject);
begin
  if Assigned(lstFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFields.Selected.Data).CharSet := CharSet;
    SetEnabledControls;
  end;
end;

procedure TfmQImport3FormatsEditor.FillFieldsList;
var
  i, j: integer;
  n: string;
  FF: TQImportFieldFormat;
  WasInActive: boolean;
begin
  WasInActive := false;
  lstFields.Items.BeginUpdate;
  try
    if lstFields.Items.Count > 0 then lstFields.Items.Clear;
    if not Assigned(DataSet) then Exit;
    WasInActive :=  not Dataset.Active;
    if WasInActive then
    try
      Dataset.Open;
    except
      Exit;
    end;
    for i := 0 to DataSet.FieldCount - 1 do
    begin
      n := DataSet.Fields[i].FieldName;
      FF := TQImportFieldFormat.Create(nil);
      FF.FieldName := n;
      with lstFields.Items.Add do
      begin
        Caption := n;
        Data := FF;
        ImageIndex := 0;
      end;
      j := FieldFormats.IndexByName(n);
      if j = -1 then Continue;
      FF.GeneratorValue := FieldFormats[j].GeneratorValue;
      FF.GeneratorStep := FieldFormats[j].GeneratorStep;
      FF.ConstantValue := FieldFormats[j].ConstantValue;
      FF.NullValue := FieldFormats[j].NullValue;
      FF.DefaultValue := FieldFormats[j].DefaultValue;
      FF.QuoteAction := FieldFormats[j].QuoteAction;
      FF.LeftQuote := FieldFormats[j].LeftQuote;
      FF.RightQuote := FieldFormats[j].RightQuote;
      FF.CharCase := FieldFormats[j].CharCase;
      FF.CharSet := FieldFormats[j].CharSet;
      FF.Replacements := FieldFormats[j].Replacements;
    end;
    if lstFields.Items.Count > 0 then lstFields.Selected := lstFields.Items[0];
  finally
    lstFields.Items.EndUpdate;
    if WasInActive then
    try
      Dataset.Close;
    except
    end;
  end;

  if lstFields.Items.Count > 0 then
  begin
    lstFields.Items[0].Focused := true;
    lstFields.Items[0].Selected := true;
    ShowFormatItem(lstFields.Items[0]);
  end;
end;

procedure TfmQImport3FormatsEditor.ApplyChanges;
var
  i, j: integer;
begin
  if lstFields.Items.Count = 0 then Exit;
  for i := 0 to lstFields.Items.Count - 1 do begin
    j := FieldFormats.IndexByName(lstFields.Items[i].Caption);
    if j > -1 then
      FieldFormats[j].Assign(TQImportFieldFormat(lstFields.Items[i].Data))
    else begin
      if not TQImportFieldFormat(lstFields.Items[i].Data).IsDefaultValues then
        with FieldFormats.Add do begin
          Assign(TQImportFieldFormat(lstFields.Items[i].Data));
        end;
    end;
  end;
end;

function TfmQImport3FormatsEditor.DataByName(const AName: string): TQImportFieldFormat;
var
  i: integer;
begin
  Result := nil;
  for i := 0 to lstFields.Items.Count - 1 do begin
    if Assigned(lstFields.Items[i].Data) then
      if AnsiCompareText(TQImportFieldFormat(lstFields.Items[i].Data).FieldName, AName) = 0 then
        Result := TQImportFieldFormat(lstFields.Items[i].Data);
  end;
end;

procedure TfmQImport3FormatsEditor.SetEnabledControls;
var
  MainCondition: boolean;
begin
  MainCondition := lstFields.Items.Count > 0;

  // generator
  edGeneratorValue.Enabled := MainCondition and
    (ConstantValue = EmptyStr) and (NullValue = EmptyStr) and
    (DefaultValue = EmptyStr) and (LeftQuote = EmptyStr) and
    (RightQuote = EmptyStr) and (QuoteAction = qaNone) and
    (CharCase = iccNone) and (CharSet = icsNone);
  laGeneratorValue.Enabled := edGeneratorValue.Enabled;
  edGeneratorStep.Enabled := edGeneratorValue.Enabled;
  laGeneratorStep.Enabled := edGeneratorValue.Enabled;
  //constant
  edConstantValue.Enabled := MainCondition and
    (GeneratorStep = 0) and (NullValue = EmptyStr) and
    (DefaultValue = EmptyStr);
  laConstantValue.Enabled := edConstantValue.Enabled;
  // default
  edNullValue.Enabled := MainCondition and
    (GeneratorStep = 0) and (ConstantValue = EmptyStr);
  laNullValue.Enabled := edNullValue.Enabled;
  edDefaultValue.Enabled := edNullValue.Enabled;
  laDefaultValue.Enabled := edNullValue.Enabled;
  // quote
  edLeftQuote.Enabled := MainCondition and (GeneratorStep = 0);
  laLeftQuote.Enabled := edLeftQuote.Enabled;
  laRightQuote.Enabled := edLeftQuote.Enabled;
  laQuoteAction.Enabled := edLeftQuote.Enabled;
  cbQuoteAction.Enabled := edLeftQuote.Enabled;
  // string coversion
  cbCharCase.Enabled := MainCondition and (GeneratorStep = 0);
  laCharCase.Enabled := cbCharCase.Enabled;
  cbCharSet.Enabled := cbCharCase.Enabled;
  laCharSet.Enabled := cbCharCase.Enabled;
  // replacements
  lvReplacements.Enabled := MainCondition;
  tbReplacements.Enabled := MainCondition;
end;

procedure TfmQImport3FormatsEditor.FormKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
var
  i: integer;
  LI: TListItem;
begin
  if (ActiveControl = lstFields) and (Shift = []) then
    case Key of
      VK_INSERT: begin
        i := lstFields.Items.Count + 1;
        while Assigned(DataByName('NewItem_' + IntToStr(i))) do Inc(i);
        LI := lstFields.Items.Add;
        LI.Caption := 'NewItem_' + IntToStr(i);
        LI.Data := TQImportFieldFormat.Create(nil);
        lstFields.Selected := LI;
        LI.EditCaption;
      end;
      VK_DELETE: begin
        LI := lstFields.Selected;
        if Assigned(LI) then
        begin
          if Assigned(LI.Data) then
            TQImportFieldFormat(LI.Data).Free;
          lstFields.Items.Delete(LI.Index);
        end;
      end;
    end;
end;

procedure TfmQImport3FormatsEditor.lstFieldsEdited(Sender: TObject;
  Item: TListItem; var S: String);
begin
  S := AnsiUpperCase(S);
  TQImportFieldFormat(Item.Data).FieldName := S;
end;

function TfmQImport3FormatsEditor.GetQImportFieldFormats: TQImportFieldFormats;
var
  PropInfo: PPropInfo;
begin
  Result := nil;
  try
    PropInfo := GetPropInfo(FComponent.ClassInfo, 'FieldFormats');
    if Assigned(PropInfo) then
      Result := TQImportFieldFormats(GetOrdProp(FComponent, PropInfo));
  except
  end;
end;

function TfmQImport3FormatsEditor.GetDataSet: TDataSet;
var
  PropInfo: PPropInfo;
begin
  Result := nil;
  try
    PropInfo := GetPropInfo(FComponent.ClassInfo, 'DataSet');
    if Assigned(PropInfo) then
      Result := TDataSet(GetOrdProp(FComponent, PropInfo));
  except
  end;
end;

procedure TfmQImport3FormatsEditor.lstFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if Change <> ctState then Exit;
  if not Assigned(Item.Data) then Exit;
  if not FIsNecessaryListChange then Exit;
  

  GeneratorValue := TQImportFieldFormat(Item.Data).GeneratorValue;
  GeneratorStep := TQImportFieldFormat(Item.Data).GeneratorStep;
  ConstantValue := TQImportFieldFormat(Item.Data).ConstantValue;
  NullValue := TQImportFieldFormat(Item.Data).NullValue;
  DefaultValue := TQImportFieldFormat(Item.Data).DefaultValue;
  LeftQuote := TQImportFieldFormat(Item.Data).LeftQuote;
  RightQuote := TQImportFieldFormat(Item.Data).RightQuote;
  QuoteAction := TQImportFieldFormat(Item.Data).QuoteAction;
  CharCase := TQImportFieldFormat(Item.Data).CharCase;
  CharSet := TQImportFieldFormat(Item.Data).CharSet;
  SetEnabledControls;
  TuneButtons;
end;

procedure TfmQImport3FormatsEditor.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  while lstFields.Items.Count > 0 do
    if Assigned(lstFields.Items[0].Data) then
    begin
      FIsNecessaryListChange := False;
      TQImportFieldFormat(lstFields.Items[0].Data).Free;
      lstFields.Items[0].Delete;
    end;
end;

procedure TfmQImport3FormatsEditor.FormCreate(Sender: TObject);
begin
  FIsNecessaryListChange := True;
  FLoadingFormatItem := False;
  FFormatItem := nil;
  TuneButtons;
end;

procedure TfmQImport3FormatsEditor.ShowFormatItem(Item: TListItem);
var
  FieldFormat: TQImportFieldFormat;
  i: integer;
begin
  FLoadingFormatItem := true;
  try
    FieldFormat := TQImportFieldFormat(Item.Data);
    edGeneratorValue.Text := IntToStr(FieldFormat.GeneratorValue);
    edGeneratorStep.Text := IntToStr(FieldFormat.GeneratorStep);
    edConstantValue.Text := FieldFormat.ConstantValue;
    edNullValue.Text := FieldFormat.NullValue;
    edDefaultValue.Text := FieldFormat.DefaultValue;
    edLeftQuote.Text := FieldFormat.LeftQuote;
    edRightQuote.Text := FieldFormat.RightQuote;
    cbQuoteAction.ItemIndex := Integer(FieldFormat.QuoteAction);
    cbCharCase.ItemIndex := Integer(FieldFormat.CharCase);
    cbCharSet.ItemIndex := Integer(FieldFormat.CharSet);

    lvReplacements.Items.BeginUpdate;
    try
      lvReplacements.Items.Clear;
      for i := 0 to FieldFormat.Replacements.Count - 1 do
        with lvReplacements.Items.Add do begin
          Caption := FieldFormat.Replacements[i].TextToFind;
          SubItems.Add(FieldFormat.Replacements[i].ReplaceWith);
          if FieldFormat.Replacements[i].IgnoreCase
            then SubItems.Add(QImportLoadStr(QIWDF_IgnoreCase_Yes))
            else SubItems.Add(QImportLoadStr(QIWDF_IgnoreCase_No));
          ImageIndex := 3;
          Data := FieldFormat.Replacements[i];
        end;
      if lvReplacements.Items.Count > 0 then begin
        lvReplacements.Items[0].Focused := true;
        lvReplacements.Items[0].Selected := true;
      end;
    finally
      lvReplacements.Items.EndUpdate;
    end;

    SetEnabledControls;
  finally
    FLoadingFormatItem := false;
  end;
end;

procedure TfmQImport3FormatsEditor.lvReplacementsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  TuneButtons;
end;

procedure TfmQImport3FormatsEditor.lvReplacementsDblClick(Sender: TObject);
begin
  if tbtEditReplacement.Enabled then
    tbtEditReplacement.Click;
end;

procedure TfmQImport3FormatsEditor.tbtAddReplacementClick(Sender: TObject);
var
  TextToFind, ReplaceWith: string;
  IgnoreCase: boolean;
  R: TQImportReplacement;
begin
  TextToFind := EmptyStr;
  ReplaceWith := EmptyStr;
  IgnoreCase := false;
  if ReplacementEdit(TextToFind, ReplaceWith, IgnoreCase) then begin
    R := TQImportFieldFormat(lstFields.ItemFocused.Data).Replacements.Add;
    R.TextToFind := TextToFind;
    R.ReplaceWith := ReplaceWith;
    R.IgnoreCase := IgnoreCase;
    with lvReplacements.Items.Add do begin
      Caption := R.TextToFind;
      SubItems.Add(R.ReplaceWith);
      if R.IgnoreCase
        then SubItems.Add(QImportLoadStr(QIWDF_IgnoreCase_Yes))
        else SubItems.Add(QImportLoadStr(QIWDF_IgnoreCase_No));
      ImageIndex := 1;
      Data := R;
      Focused := true;
      Selected := true;
    end;
  end;
  TuneButtons;
end;

procedure TfmQImport3FormatsEditor.tbtEditReplacementClick(Sender: TObject);
var
  TextToFind, ReplaceWith: string;
  IgnoreCase: boolean;
  R: TQImportReplacement;
begin
  R := TQImportReplacement(lvReplacements.ItemFocused.Data);
  TextToFind := R.TextToFind;
  ReplaceWith := R.ReplaceWith;
  IgnoreCase := R.IgnoreCase;
  if ReplacementEdit(TextToFind, ReplaceWith, IgnoreCase) then begin
    R.TextToFind := TextToFind;
    R.ReplaceWith := ReplaceWith;
    R.IgnoreCase := IgnoreCase;
    with lvReplacements.ItemFocused do begin
      Caption := R.TextToFind;
      SubItems[0] := R.ReplaceWith;
      if R.IgnoreCase
        then SubItems[1] := QImportLoadStr(QIWDF_IgnoreCase_Yes)
        else SubItems[1] := QImportLoadStr(QIWDF_IgnoreCase_No);
    end;
  end;
  TuneButtons;
end;

procedure TfmQImport3FormatsEditor.tbtDelReplacementClick(Sender: TObject);
begin
  TQImportReplacement(lvReplacements.ItemFocused.Data).Free;
  lvReplacements.ItemFocused.Delete;
  TuneButtons;
end;

procedure TfmQImport3FormatsEditor.TuneButtons;
begin
  tbtAddReplacement.Enabled := Assigned(lstFields.ItemFocused);
  tbtEditReplacement.Enabled := Assigned(lstFields.ItemFocused) and
                                Assigned(lvReplacements.ItemFocused);
  tbtDelReplacement.Enabled := Assigned(lstFields.ItemFocused) and
                               Assigned(lvReplacements.ItemFocused);
end;

end.
