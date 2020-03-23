unit fuQImport3ProgressDlg;
{$I QImport3VerCtrl.Inc}
interface

uses
  {$IFDEF VCL16}
    Winapi.Windows,
    Winapi.Messages,
    System.SysUtils,
    System.Classes,
    Vcl.Controls,
    Vcl.Forms,
    Vcl.StdCtrls,
    Vcl.ExtCtrls,
    Vcl.ComCtrls,
  {$ELSE}
    Windows,
    Messages,
    SysUtils,
    Classes,
    Controls,
    Forms,
    ExtCtrls,
    StdCtrls,
    ComCtrls,
  {$ENDIF}
  QImport3,
  QImport3StrTypes;

const
  WM_QIMPORT_PROGRESS = WM_USER + 1;

  QIP_STATE = 1;
  QIP_ERROR = 2;
  QIP_ERROR_ADV = 7;
  QIP_IMPORT = 3;
  QIP_COMMIT = 4;
  QIP_FINISH = 5;
  QIP_ROWCOUNT = 6;

type
  TImportState = (istPrepare, istImport, istCancel, istFinish, istPause,
   istContinue);

  TfmQImport3ProgressDlg = class(TForm)
    bCancel: TButton;
    paInfo: TPanel;
    bvProcessed: TBevel;
    prbImport: TProgressBar;
    laProcessed: TLabel;
    bvStateValue: TBevel;
    laProcessedValue: TLabel;
    bvState: TBevel;
    bvProcessedValue: TBevel;
    laState: TLabel;
    laStateValue: TLabel;
    bvTime: TBevel;
    laTime: TLabel;
    laTimeValue: TLabel;
    bvTimeValue: TBevel;
    Timer: TTimer;
    laPercent: TLabel;
    laCommitted: TLabel;
    bvCommitted: TBevel;
    laCommittedValue: TLabel;
    bvCommittedValue: TBevel;
    bvInserted: TBevel;
    bvUpdates: TBevel;
    bvDeleted: TBevel;
    bvErrors: TBevel;
    laInserted: TLabel;
    laUpdated: TLabel;
    laDeleted: TLabel;
    laErrors: TLabel;
    bvInsertedValue: TBevel;
    laInsertedValue: TLabel;
    bvUpdatedValue: TBevel;
    laUpdatedValue: TLabel;
    laDeletedValue: TLabel;
    bvDeletedValue: TBevel;
    laErrorsValue: TLabel;
    bvErrorsValue: TBevel;
    bvPervent: TBevel;
    procedure TimerTimer(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bCancelClick(Sender: TObject);
  private
    FImportState: TImportState;

    FInsertedValue: integer;
    FUpdatedValue: integer;
    FDeletedValue: integer;
    FSkippedValue: integer;

    FErrorsValue: integer;
    FAdvErrorsValue: Integer;
    FRowCount: integer;
    FTime: TDateTime;
    FErrors: TqiStrings;
    FImport: TQImport3;

    procedure SetCaptions;

    procedure SetImportState(const Value: TImportState);

    procedure SetInsertedValue(const Value: integer);
    procedure SetUpdatedValue(const Value: integer);
    procedure SetDeletedValue(const Value: integer);
    procedure SetSkippedValue(const Value: integer);

    procedure SetErrorsValue(const Value: integer);
    procedure SetAdvErrorsValue(const Value: Integer);
    procedure SetRowCount(const Value: integer);
    procedure SetTime(const ATime: TDateTime);

    procedure ShowImportedState;

    procedure ShowProcessedValue;
    procedure ShowInsertedValue;
    procedure ShowUpdatedValue;
    procedure ShowDeletedValue;
    procedure ShowErrorsValue;

    procedure ShowTime;
    procedure ShowPercent;

    function GetStateString: string;
    procedure SetVisibleProgress(const Value: boolean);
  protected
    procedure CreateParams( var Params: TCreateParams ); override;
    procedure ImportProgress(var Msg: TMessage); message WM_QIMPORT_PROGRESS;
  public
    constructor CreateProgress(AOwner: TComponent; AImport: TQImport3);
    property ImportState: TImportState read FImportState write SetImportState;
    property StateString: string read GetStateString;

    property InsertedValue: integer read FInsertedValue write SetInsertedValue;
    property UpdatedValue: integer read FUpdatedValue write SetUpdatedValue;
    property DeletedValue: integer read FDeletedValue write SetDeletedValue;
    property SkippedValue: integer read FSkippedValue write SetSkippedValue;

    property ErrorsValue: integer read FErrorsValue write SetErrorsValue;
    property AdvErrorsValue: Integer read FAdvErrorsValue write SetAdvErrorsValue;
    property RowCount: integer read FRowCount write SetRowCount;
    property Errors: TqiStrings read FErrors write FErrors;
    property Import: TQImport3 read FImport write FImport;

    procedure Stop(ShowErrors: boolean);
    procedure StepIt;
    procedure Start;
    procedure ShowCommittedValue;
  end;


implementation

{$R *.DFM}

uses
  QImport3StrIDs;

{ TfmQImport3ProgressDlg }

constructor TfmQImport3ProgressDlg.CreateProgress(AOwner: TComponent;
  AImport: TQImport3);
begin
  inherited Create(AOwner);
  FImport := AImport;
  FImportState := istPrepare;

  FInsertedValue := 0;
  FUpdatedValue := 0;
  FDeletedValue := 0;
  FSkippedValue := 0;

  FErrorsValue := 0;
  FAdvErrorsValue := 0;
  FRowCount := 0;
  FTime := 0;
  prbImport.Max := 0;
  SetVisibleProgress(false);
end;

procedure TfmQImport3ProgressDlg.CreateParams(var Params: TCreateParams);
begin
  inherited;
  if Assigned(Owner) and (Owner is TForm) then
    with Params do
    begin
      Style := Style or ws_Overlapped;
      WndParent := (Owner as TForm).Handle;
    end;
end;

procedure TfmQImport3ProgressDlg.ImportProgress(var Msg: TMessage);
begin
  if Msg.Msg = WM_QIMPORT_PROGRESS then
  begin
    case Msg.WParam of
      QIP_STATE:
        begin
          ImportState := TImportState(Msg.LParam);
          if ImportState in [istPrepare, istImport] then
            Start;
          Timer.Enabled := not (ImportState in [istPause, istCancel, istFinish]);
        end;
      QIP_ERROR:
        begin
          ErrorsValue := ErrorsValue + Msg.LParam;
          StepIt;
        end;
      QIP_IMPORT:
        begin
          case TQImportAction(Msg.LParam) of
            qiaInsert: InsertedValue := InsertedValue + 1;
            qiaUpdate: UpdatedValue := UpdatedValue + 1;
            qiaDelete: DeletedValue := DeletedValue + 1;
            qiaNone: SkippedValue := SkippedValue + 1;
          end;
          StepIt;
        end;
      QIP_COMMIT: ShowCommittedValue;
      QIP_FINISH: Stop(Boolean(Msg.LParam));
      QIP_ROWCOUNT: RowCount := Msg.LParam;
      QIP_ERROR_ADV: AdvErrorsValue := AdvErrorsValue + Msg.LParam;
    end;
  end;
  inherited;
end;

procedure TfmQImport3ProgressDlg.SetCaptions;
begin
  laState.Caption := QImportLoadStr(QIPD_State);
  laProcessed.Caption := QImportLoadStr(QIPD_Processed);
  laInserted.Caption := QImportLoadStr(QIPD_Inserted);
  laUpdated.Caption := QImportLoadStr(QIPD_Updated);
  laDeleted.Caption := QImportLoadStr(QIPD_Deleted);
  laErrors.Caption := QImportLoadStr(QIPD_Errors);
  laCommitted.Caption := QImportLoadStr(QIPD_Committed);
  laTime.Caption := QImportLoadStr(QIPD_Time);
  bCancel.Caption := QImportLoadStr(QIPD_Cancel);

end;

procedure TfmQImport3ProgressDlg.FormShow(Sender: TObject);
begin
  SetCaptions;

  ShowImportedState;

  ShowProcessedValue;
  ShowInsertedValue;
  ShowUpdatedValue;
  ShowDeletedValue;

  ShowCommittedValue;
  ShowErrorsValue;

  ShowTime;
end;

procedure TfmQImport3ProgressDlg.TimerTimer(Sender: TObject);
begin
  SetTime(FTime + 0.00001);
end;

procedure TfmQImport3ProgressDlg.bCancelClick(Sender: TObject);
begin
  Import.Cancel;
end;

// SET & GET //

procedure TfmQImport3ProgressDlg.SetImportState(const Value: TImportState);
begin
  if FImportState <> Value then
  begin
    FImportState := Value;
    ShowImportedState;
  end;
end;

procedure TfmQImport3ProgressDlg.SetInsertedValue(const Value: integer);
begin
  if FInsertedValue <> Value then
  begin
    FInsertedValue := Value;
    ShowProcessedValue;
    ShowInsertedValue;
  end;
end;

procedure TfmQImport3ProgressDlg.SetUpdatedValue(const Value: integer);
begin
  if FUpdatedValue <> Value then
  begin
    FUpdatedValue := Value;
    ShowProcessedValue;
    ShowUpdatedValue;
  end;
end;

procedure TfmQImport3ProgressDlg.SetDeletedValue(const Value: integer);
begin
  if FDeletedValue <> Value then
  begin
    FDeletedValue := Value;
    ShowProcessedValue;
    ShowDeletedValue;
  end;
end;

procedure TfmQImport3ProgressDlg.SetSkippedValue(const Value: integer);
begin
  if FSkippedValue <> Value then
  begin
    FSkippedValue := Value;
    ShowProcessedValue;
  end;
end;

procedure TfmQImport3ProgressDlg.SetErrorsValue(const Value: integer);
begin
  if FErrorsValue <> Value then
  begin
    FErrorsValue := Value;
    ShowErrorsValue;
    ShowProcessedValue;
  end;
end;

procedure TfmQImport3ProgressDlg.SetRowCount(const Value: integer);
begin
  if FRowCount <> Value then
  begin
    FRowCount := Value;
    SetVisibleProgress(FRowCount <> 0);
    prbImport.Max := FRowCount;
  end;
end;

procedure TfmQImport3ProgressDlg.SetTime(const ATime: TDateTime);
begin
  if FTime <> ATime then
  begin
    FTime := ATime;
    ShowTime;
  end;
end;

function TfmQImport3ProgressDlg.GetStateString: string;
begin
  case FImportState of
    istPrepare: Result := QImportLoadStr(QIPD_Preparing);//'Preparing...';
    istImport,
    istContinue : Result := QImportLoadStr(QIPD_Importing);
    istCancel : Result := QImportLoadStr(QIPD_Aborted);
    istFinish : Result := QImportLoadStr(QIPD_Finished);
    istPause: Result := QImportLoadStr(QIPD_Paused);
  else Result := '';
  end;
end;

procedure TfmQImport3ProgressDlg.SetVisibleProgress(const Value: boolean);
begin
  laPercent.Enabled := Value;
  laPercent.Visible := Value;
  prbImport.Enabled := Value;
  prbImport.Visible := Value;
end;

// SHOW //

procedure TfmQImport3ProgressDlg.ShowImportedState;
begin
  laStateValue.Caption := StateString;
end;

procedure TfmQImport3ProgressDlg.ShowProcessedValue;
begin
  laProcessedValue.Caption := IntToStr(FInsertedValue + FUpdatedValue +
    FDeletedValue + FSkippedValue + FErrorsValue);
end;

procedure TfmQImport3ProgressDlg.ShowInsertedValue;
begin
  laInsertedValue.Caption := IntToStr(FInsertedValue);
end;

procedure TfmQImport3ProgressDlg.ShowUpdatedValue;
begin
  laUpdatedValue.Caption := IntToStr(FUpdatedValue);
end;

procedure TfmQImport3ProgressDlg.ShowDeletedValue;
begin
  laDeletedValue.Caption := IntToStr(FDeletedValue);
end;

procedure TfmQImport3ProgressDlg.ShowCommittedValue;
begin
  laCommittedValue.Caption :=
    IntToStr(InsertedValue + UpdatedValue + DeletedValue);
end;

procedure TfmQImport3ProgressDlg.ShowErrorsValue;
begin
  laErrorsValue.Caption := IntToStr(FErrorsValue + FAdvErrorsValue);
end;

procedure TfmQImport3ProgressDlg.ShowTime;
begin
  laTimeValue.Caption := FormatDateTime('h:nn:ss', FTime);
end;

procedure TfmQImport3ProgressDlg.ShowPercent;
begin
  if (prbImport.Position <> 0) and (prbImport.Max <> 0) then begin
    laPercent.Caption := IntToStr(Round(prbImport.Position * 100 / prbImport.Max)) + '%';
  end
  else laPercent.Caption := '0%';
end;

procedure TfmQImport3ProgressDlg.Start;
begin
  bCancel.Caption := QImportLoadStr(QIPD_Cancel);
  bCancel.ModalResult := mrCancel;
  bCancel.OnClick := bCancelClick;
  Caption := QImportLoadStr(QIPD_Importing); // Importing...
  Timer.Enabled := true;
end;

procedure TfmQImport3ProgressDlg.Stop(ShowErrors: boolean);
begin
  bCancel.Caption := QImportLoadStr(QIPD_OK);
  bCancel.ModalResult := mrOk;
  bCancel.OnClick := nil;
  Caption := QImportLoadStr(QIPD_Import_finished);
  Timer.Enabled := false;
end;

procedure TfmQImport3ProgressDlg.StepIt;
begin
  prbImport.StepIt;
  ShowPercent;
end;

procedure TfmQImport3ProgressDlg.SetAdvErrorsValue(const Value: Integer);
begin
  if FAdvErrorsValue <> Value then
  begin
    FAdvErrorsValue := Value;
    ShowErrorsValue;
    ShowProcessedValue;
  end;
end;

end.
