unit QImport3XLS;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    System.Classes,
    Data.Db,
    System.IniFiles,
  {$ELSE}
    Classes,
    Db,
    IniFiles,
  {$ENDIF}
  QImport3,
  QImport3XLSFile,
  QImport3XLSMapParser;

type
  TQImportRange = (qirMax, qirMin);
  TQImportSheetIDType = (qstNumber, qstName);

  TQImport3XLS = class(TQImport3)
  private
    FXLS: TXLSFile;
    FMapRowList: TMapRowList;

    FMax: integer;
    FMin: integer;

    FImportRange: TQImportRange;
    FDefaultSheetIDType: TQImportSheetIDType;
    FDefaultSheetNumber: integer;
    FDefaultSheetName: string;

    FRow: TStrings;
    //---
    FTotal: integer;
    FCounter: integer;
    //---
  protected
    procedure BeforeImport; override;
    procedure AfterImport; override;

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
    destructor Destroy; override;
  published
    property FileName;
    property SkipFirstRows default 0;
    property SkipFirstCols default 0;
    property ImportRange: TQImportRange read FImportRange
      write FImportRange default qirMax;
    property DefaultSheetIDType: TQImportSheetIDType read FDefaultSheetIDType
      write FDefaultSheetIDType default qstNumber;
    property DefaultSheetNumber: integer read FDefaultSheetNumber
      write FDefaultSheetNumber default 1;
    property DefaultSheetName: string read FDefaultSheetName
      write FDefaultSheetName;
  end;

implementation

uses
  {$IFDEF VCL16}
    System.Math,
    System.SysUtils,
  {$ELSE}
    Math,
    SysUtils,
  {$ENDIF}
  QImport3Common;

{ TQImport3XLS }

constructor TQImport3XLS.Create(AOwner: TComponent);
begin
  inherited;
  FImportRange := qirMax;
  FDefaultSheetIDType := qstNumber;
  FDefaultSheetNumber := 1;
  FDefaultSheetName := EmptyStr;

  SkipFirstRows := 0;
  SkipFirstCols := 0;

  FRow := TStringList.Create;
end;

destructor TQImport3XLS.Destroy;
begin
  FRow.Free;
  inherited;
end;

procedure TQImport3XLS.BeforeImport;
var
  i: integer;
  MapRow: TMapRow;
begin
  FXLS := TXLSFile.Create;
  FXLS.FileName := FileName;
  FXLS.Load;
  FMapRowList := TMapRowList.Create(FXLS);
  FMapRowList.SkipFirstRows := SkipFirstRows;
  FMapRowList.SkipFirstCols := SkipFirstCols;
  for i := 0 to Map.Count - 1 do
  begin
    MapRow := TMapRow.Create(FMapRowList);
    try
      ParseMapString(Map.Values[Map.Names[i]], MapRow);
      MapRow.Update;
    except
      MapRow.Free;
      raise;
    end;
    FMapRowList.Add(MapRow);
  end;
  FMapRowList.Update;

  FMin := FMapRowList.Items[FMapRowList.MinRow].Length;
  FMax := FMapRowList.Items[FMapRowList.MaxRow].Length;
  case FImportRange of
    qirMax: FTotalRecCount := FMax;
    qirMin: FTotalRecCount := FMin;
  end;
  inherited BeforeImport;
end;

procedure TQImport3XLS.StartImport;
begin
  FTotal := 0;
  case FImportRange of
    qirMax: FTotal := FMax - 1;
    qirMin: FTotal := FMin - 1;
  end;
  if ImportRecCount > 0 then
    {$IFDEF VCL15}
    FTotal := Trunc(MinIntValue([FTotal, ImportRecCount - 1]));
    {$ELSE}
    FTotal := Trunc(MinValue([FTotal, ImportRecCount - 1]));
    {$ENDIF}
  FCounter := 0;

  CurrentLineNumber := CurrentLineNumber + SkipFirstRows;
end;

function TQImport3XLS.CheckCondition: boolean;
begin
  Result := FCounter <= FTotal;
end;

function TQImport3XLS.Skip: boolean;
begin
  Result := false;
end;

procedure TQImport3XLS.FillImportRow;
var
  j, k: integer;
  wstr: WideString;
  p: Pointer;
begin
  FImportRow.ClearValues;
  RowIsEmpty := True;
  for j := 0 to FImportRow.Count - 1 do
  begin
    if FImportRow.MapNameIdxHash.Search(FImportRow[j].Name, p) then
    begin
      k := Integer(p);
      wstr := FMapRowList[k].GetCellValue(FCounter + 1);
      if AutoTrimValue then
        wstr := Trim(wstr);
      RowIsEmpty := RowIsEmpty and (wstr = '');
      FImportRow.SetValue(Map.Names[k], wstr, False);
    end;
    DoUserDataFormat(FImportRow[j]);
  end;
end;

function TQImport3XLS.ImportData: TQImportResult;
begin
  Result := qirOk;
  try
    try
      if Canceled and not CanContinue then
      begin
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

procedure TQImport3XLS.ChangeCondition;
begin
  Inc(FCounter)
end;

procedure TQImport3XLS.FinishImport;
begin
  if not Canceled and not IsCSV then
  begin
    if CommitAfterDone then
      DoNeedCommit
    else if (CommitRecCount > 0) and ((ImportedRecs + ErrorRecs) mod CommitRecCount > 0) then
      DoNeedCommit;
  end;    
end;

procedure TQImport3XLS.AfterImport;
begin
  inherited;
  FMapRowList.Free;
  if Assigned(FXLS) then begin
    FXLS.Free;
    FXLS := nil;
  end;
end;

procedure TQImport3XLS.DoLoadConfiguration(IniFile: TIniFile);
begin
  inherited;
   with IniFile do begin
     SkipFirstCols := ReadInteger(XLS_OPTIONS, XLS_SKIP_COLS, SkipFirstCols);
     SkipFirstRows := ReadInteger(XLS_OPTIONS, XLS_SKIP_ROWS, SkipFirstRows);
   end;
end;

procedure TQImport3XLS.DoSaveConfiguration(IniFile: TIniFile);
begin
  inherited;
   with IniFile do begin
     WriteInteger(XLS_OPTIONS, XLS_SKIP_COLS, SkipFirstCols);
     WriteInteger(XLS_OPTIONS, XLS_SKIP_ROWS, SkipFirstRows);
   end;
end;

end.
