unit QImport3CustomControl;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF QI_UNICODE}

uses
  {$IFDEF VCL16}
    Vcl.Controls,
    Winapi.Windows,
    Winapi.Messages,
    System.Classes,
  {$ELSE}
    Controls,
    Windows,
    Messages,
    Classes,
  {$ENDIF}
  QImport3WideStringCanvas;

type
  TEmsCustomControl = class(TWinControl)
  private
    FCanvas: TEmsWideStringCanvas;
    procedure WMPaint(var Message: TWMPaint); message WM_PAINT;
  protected
    procedure Paint; virtual;
    procedure PaintWindow(DC: HDC); override;
    property Canvas: TEmsWideStringCanvas read FCanvas;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

{$ENDIF}

implementation

{$IFDEF QI_UNICODE}

{ TEmsCustomControl }

constructor TEmsCustomControl.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FCanvas := TEmsWideStringCanvas.Create;
  TEmsWideStringCanvas(FCanvas).Control := Self;
end;

destructor TEmsCustomControl.Destroy;
begin
  FCanvas.Free;
  inherited Destroy;
end;

procedure TEmsCustomControl.Paint;
begin
end;

procedure TEmsCustomControl.PaintWindow(DC: HDC);
begin
  FCanvas.Lock;
  try
    FCanvas.Handle := DC;
    try
      TEmsWideStringCanvas(FCanvas).UpdateTextFlags;
      Paint;
    finally
      FCanvas.Handle := 0;
    end;
  finally
    FCanvas.Unlock;
  end;
end;

procedure TEmsCustomControl.WMPaint(var Message: TWMPaint);
begin
  ControlState := ControlState + [csCustomPaint];
  inherited;
  ControlState := ControlState - [csCustomPaint];
end;

{$ENDIF}

end.
