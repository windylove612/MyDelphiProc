unit QImport3DataSet;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    Data.Db,
    System.Classes,
    System.IniFiles,
  {$ELSE}
    Db,
    Classes,
    IniFiles,
  {$ENDIF}
  QImport3;

type
  TQImport3DataSet = class(TQImport3)
  private
    FSource: TDataSet;
    FGoToFirstRecord: boolean;
    //---
    FSkipCounter: integer;
    FImportCounter: integer;
    //---
  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    function CheckProperties: boolean; override;

    procedure StartImport; override;
    function CheckCondition: boolean; override;
    function Skip: boolean; override;
    procedure FillImportRow; override;
    function ImportData: TQImportResult; override;
    procedure ChangeCondition; override;
    procedure FinishImport; override;

    procedure DoLoadConfiguration(IniFile: TIniFile); override;
    procedure DoSaveConfiguration(IniFile: TIniFile); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property Source: TDataSet read FSource write FSource;
    property SkipFirstRows default 0;
    property GoToFirstRecord: boolean read FGoToFirstRecord
      write FGoToFirstRecord default true;
  end;

implementation

uses
  {$IFDEF VCL16}
    System.SysUtils,
  {$ELSE}
    SysUtils,
  {$ENDIF}
  QImport3StrIDs,
  QImport3Common, 
  QImport3StrTypes;

{ TQImport3DataSet }

constructor TQImport3DataSet.Create(AOwner: TComponent);
begin
  inherited;
  SkipFirstRows := 0;
  FGoToFirstRecord := true;
end;

procedure TQImport3DataSet.DoLoadConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do begin
    SkipFirstRows := ReadInteger(DATA_SET_OPTIONS, DATA_SET_SKIP_LINES, SkipFirstRows);
  end;
end;

procedure TQImport3DataSet.DoSaveConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do begin
    WriteInteger(DATA_SET_OPTIONS, DATA_SET_SKIP_LINES, SkipFirstRows);
  end;
end;

procedure TQImport3DataSet.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited;
  if (AComponent = FSource) and (Operation = opRemove) then FSource := nil;
end;

procedure TQImport3DataSet.StartImport;
begin
  Source.DisableControls;
  if FGoToFirstRecord then Source.First;
  FSkipCounter := 0;
  FImportCounter := 0;
end;

function TQImport3DataSet.CheckCondition: boolean;
begin
  Result := not Source.Eof and
            ((ImportRecCount = 0) or (FImportCounter < ImportRecCount));
end;

function TQImport3DataSet.Skip: boolean;
begin
  Result := not Source.Eof and (FSkipCounter < SkipFirstRows);
  Inc(FSkipCounter);
end;

procedure TQImport3DataSet.FillImportRow;
var
  i, k: integer;
  DField, SField: TField;
  p: Pointer;
  StrValue: qiString;
begin
  FImportRow.ClearValues;
  RowIsEmpty := True;
  for i := 0 to FImportRow.Count - 1 do
  begin
    if FImportRow.MapNameIdxHash.Search(FImportRow[i].Name, p) then
    begin
      k := Integer(p);
      DField := DataSet.FindField(Map.Names[k]);
{$IFDEF VCL7}
      SField := Source.FindField(Map.ValueFromIndex[k]);
{$ELSE}
      SField := Source.FindField(Map.Values[Map.Names[k]]);
{$ENDIF}
      if Assigned(DField) and Assigned(SField) then
      begin
        StrValue := SField.AsString;
        if AutoTrimValue then
          StrValue := Trim(StrValue);
        RowIsEmpty := RowIsEmpty and (StrValue = '');
        FImportRow.SetValue(Map.Names[k], StrValue, SField.IsBlob);
      end;
    end;
    DoUserDataFormat(FImportRow[i]);
  end;
end;

function TQImport3DataSet.ImportData: TQImportResult;
begin
  Result := qirOk;
  try
    try
      if Canceled  and not CanContinue then begin
        Result := qirBreak;
        Exit;
      end;

      DataManipulation;

    except
      on E:Exception do begin
        try
          DestinationCancel;
        except
        end;
        DoImportError(E);
        Result := qirContinue;
        Exit;
      end;
    end;
  finally
    if (not IsCSV) and (CommitRecCount > 0) and not CommitAfterDone and
       (
        ((ImportedRecs + ErrorRecs) > 0)
        and ((ImportedRecs + ErrorRecs) mod CommitRecCount = 0)
       )
    then
      DoNeedCommit;
  end;
end;

procedure TQImport3DataSet.ChangeCondition;
begin
  Source.Next;
  Inc(FImportCounter);
end;

procedure TQImport3DataSet.FinishImport;
begin
  try
    if not Canceled then
    begin
      if CommitAfterDone then
        DoNeedCommit
      else if (CommitRecCount > 0) and ((ImportedRecs + ErrorRecs) mod CommitRecCount > 0) then
        DoNeedCommit;
    end;
  finally
    Source.EnableControls;
  end;
end;

function TQImport3DataSet.CheckProperties: boolean;
var
  i: integer;
begin
  Result := false;
  if not Assigned(DataSet) then
    raise EQImportError.Create(QImportLoadStr(QIE_NoDataSet));
  if not Assigned(Source) then
    raise EQImportError.Create(QImportLoadStr(QIE_NoSource));
  if Map.Count = 0 then
    raise EQImportError.Create(QImportLoadStr(QIE_MappingEmpty));

  for i := 0 to Map.Count - 1 do begin
    if not Assigned(DataSet.FindField(Map.Names[i])) then
      raise EQImportError.CreateFmt(QImportLoadStr(QIE_FieldNotFound), [Map.Names[i]]);
    if not Assigned(Source.FindField(Map.Values[Map.Names[i]])) then
      raise EQImportError.CreateFmt(QImportLoadStr(QIE_SourceFieldNotFound),
                                    [Map.Values[Map.Names[i]]]);
  end;
  if not Result then Result := true;
end;

end.
