unit QImport3InfoPanel;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    Winapi.Windows,
    Winapi.Messages,
    System.SysUtils,
    System.Classes,
    Vcl.Controls,
    Vcl.ExtCtrls,
    Vcl.Graphics;
  {$ELSE}
    Windows,
    Messages,
    SysUtils,
    Classes,
    Controls,
    ExtCtrls,
    Graphics;
  {$ENDIF}

type
  TInfoPanel = class(TCustomControl)
  private
    FBmp: TBitmap;
    FAutoSize: Boolean;
    FShadow: Boolean;
    FWordWrap: Boolean;
    FBkColor: TColor;
    FGapText: Integer;
    FGapBitmap: Integer;
    FRadius: Integer;
    FGapVertical: Integer;
    FShadowDepth: Integer;
    function GetTransparent: Boolean;
    procedure SetTransparent(Value: Boolean);
    procedure SetShadow(Value: Boolean);
    procedure SetWordWrap(Value: Boolean);
    procedure SetBkgColor(Value: TColor);
    procedure SetGapText(Value: Integer);
    procedure SetGapBitmap(Value: Integer);
    procedure SetRadius(const Value: Integer);
    procedure SetGapVertical(Value: Integer);
    procedure SetShadowDepth(const Value: Integer);
    procedure CMFontChanged(var Message: TMessage); message CM_FONTCHANGED;
    procedure CMTextChanged(var Message: TMessage); message CM_TEXTCHANGED;
    procedure WMSize(var Message: TMessage); message WM_SIZE;
  protected
    function GetPanelText: string;
    function GetTextRect: TRect; dynamic;
    procedure AdjustBounds; dynamic;
    procedure Paint; override;
{$IFDEF VCL4}
    procedure SetEnabled(Value: Boolean); override;
{$ENDIF}
    procedure DoDrawText(var Rect: TRect; Flags: Longint);
    procedure SetAutoSize(Value: Boolean); {$IFDEF VCL6} override; {$ENDIF}
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property AutoSize: Boolean read FAutoSize write SetAutoSize default True;
    property BkColor: TColor read FBkColor write SetBkgColor default clInfoBk;
    property GapBitmap: Integer read FGapBitmap write SetGapBitmap default 5;
    property GapText: Integer read FGapText write SetGapText default 2;
    property GapVertical: Integer read FGapVertical write SetGapVertical default 6;
    property Radius: Integer read FRadius write SetRadius default 10;
    property Shadow: Boolean read FShadow write SetShadow default False;
    property ShadowDepth: Integer read FShadowDepth write SetShadowDepth default 3;
    property Transparent: Boolean read GetTransparent write SetTransparent
      default False;
    property WordWrap: Boolean read FWordWrap write SetWordWrap default False;
    { Inherited properties }
    property Align;
    {$IFDEF VCL4}
    property Anchors;
    {$ENDIF}
    property Caption;
    property Enabled;
    property Color;
    property Font;
    property Visible;
  end;

implementation

{$IFDEF VCL6}
uses
  {$IFDEF VCL16}
    System.Types;
  {$ELSE}
    Types;
  {$ENDIF}
{$ENDIF}

{$R QImport3InfoPanel.res}

{ TInfoPanel }

procedure TInfoPanel.AdjustBounds;
var
  H, W: Integer;
  TR: TRect;
begin
  if not (csReading in ComponentState) and FAutoSize then
  begin
    Canvas.Font := Self.Font;

    if Text = '' then Text := ' ';

    TR := GetTextRect;

    if Align = alNone then
      W := FGapBitmap * 2 + FBmp.Width + FGapText + (TR.Right - TR.Left)
    else
      W := Width;

    H := (TR.Bottom - TR.Top) + FGapVertical * 2;
    SetBounds(Left, Top, W, H);
  end;
end;

procedure TInfoPanel.WMSize(var Message: TMessage);
begin
  if Align <> alNone then AdjustBounds;
end;

procedure TInfoPanel.CMFontChanged(var Message: TMessage);
begin
  AdjustBounds;
  inherited;
end;

procedure TInfoPanel.CMTextChanged(var Message: TMessage);
begin
  AdjustBounds;
  Invalidate;
  inherited;
end;

constructor TInfoPanel.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  ControlStyle := ControlStyle + [csOpaque, csReplicatable];
  
  FBmp := TBitmap.Create;
  FBmp.LoadFromResourceName(HInstance, 'ENBL');
  FBmp.TransparentColor := clBlue;
  
  FAutoSize := True;
  FBkColor := clInfoBk;
  FGapBitmap := 5;
  FGapText := 2;
  FGapVertical := 6;
  FRadius := 10;
  FShadowDepth := 3;
  FShadow := False;
  FWordWrap := False;
end;

destructor TInfoPanel.Destroy;
begin
  FBmp.Free;
  inherited Destroy;
end;

procedure TInfoPanel.DoDrawText(var Rect: TRect; Flags: Integer);
var
  Text: string;
begin
  Text := GetPanelText;
  {$IFDEF VCL4}
  Flags := DrawTextBiDiModeFlags(Flags);
  {$ENDIF}
  SetBkColor(Canvas.Handle, ColorToRGB(BkColor));
  DrawText(Canvas.Handle, PChar(Text), Length(Text), Rect, Flags);
end;

function TInfoPanel.GetPanelText: string;
begin
  Result := Text; 
end;

function TInfoPanel.GetTextRect: TRect;
const
  WordWraps: array[Boolean] of Word = (0, DT_WORDBREAK);
var
  DC: HDC;
begin
  Result := Rect(0, 0, Width, Height);
  Result.Left := Result.Left + FGapBitmap + FBmp.Width + FGapText;
  DC := GetDC(0);
  Canvas.Handle := DC;
  DoDrawText(Result, (DT_EXPANDTABS or DT_CALCRECT) or WordWraps[FWordWrap]);
  Canvas.Handle := 0;
  ReleaseDC(0, DC);
end;

function TInfoPanel.GetTransparent: Boolean;
begin
  Result := not (csOpaque in ControlStyle);
end;

procedure TInfoPanel.Paint;
const
  WordWraps: array[Boolean] of Word = (0, DT_WORDBREAK);
var
  H, Y: Integer;
  SR, TR: TRect;
  OldBrushStyle: TBrushStyle;
  DrawStyle: Longint;
begin
  if not Transparent then
  begin
    Canvas.Brush.Color := Self.Color;
    Canvas.Brush.Style := bsSolid;
    Canvas.FillRect(ClientRect);
  end;

  Canvas.Brush.Style := bsClear;
  
  FBmp.Dormant;
  if Enabled then
  begin
    FBmp.LoadFromResourceName(HInstance, 'ENBL');
    Canvas.Font.Color := Self.Font.Color;
    Canvas.Pen.Color := clBlack;
  end
  else begin
    FBmp.LoadFromResourceName(HInstance, 'DSBL');
    Canvas.Font.Color := clGrayText;
    Canvas.Pen.Color := clGrayText;
  end;

  TR := GetTextRect;
  H := TR.Bottom - TR.Top;

  if Shadow then
  begin
    Canvas.Pen.Color := clGray;
    Canvas.Brush.Color := clGray;
    Canvas.RoundRect(ShadowDepth, ShadowDepth, Width, Height, Radius, Radius);
    Canvas.Pen.Color := clBlack;
    Canvas.Brush.Color := BkColor;
    Canvas.RoundRect(0, 0, Width - ShadowDepth, Height - ShadowDepth, Radius, Radius);
    TR.Top := (Height - H) div 2 - ShadowDepth div 2;
  end
  else begin
    Canvas.Brush.Color := BkColor;
    Canvas.RoundRect(0, 0, Width, Height, Radius, Radius);
    TR.Top := (Height - H) div 2;
  end;

  Y := TR.Top - 2;
  SR := Rect(GapBitmap, Y, FBmp.Width + GapBitmap, FBmp.Height + Y);
  OldBrushStyle := Canvas.Brush.Style;
  Canvas.Brush.Style := bsClear;
  Canvas.BrushCopy(SR, FBmp, Rect(0, 0, FBmp.Width, FBmp.Height), clBlue);
  Canvas.Brush.Style := OldBrushStyle;

  DrawStyle := DT_NOPREFIX or
               DT_EXPANDTABS or
               DT_LEFT or
               WordWraps[FWordWrap];

  TR.Bottom := TR.Top + H;
  DoDrawText(TR, DrawStyle);
end;

procedure TInfoPanel.SetAutoSize(Value: Boolean);
begin
  if FAutoSize <> Value then
  begin
    FAutoSize := Value;
    AdjustBounds;
    Invalidate;
  end;
end;

procedure TInfoPanel.SetBkgColor(Value: TColor);
begin
  if FBkColor <> Value then
  begin
    FBkColor := Value;
    Invalidate;
  end;
end;

procedure TInfoPanel.SetGapBitmap(Value: Integer);
begin
  if (FGapBitmap <> Value) and (Value >= 0) then
  begin
    FGapBitmap := Value;
    AdjustBounds;
    Invalidate;
  end;  
end;

procedure TInfoPanel.SetGapText(Value: Integer);
begin
  if (FGapText <> Value) and (Value >= 0) then
  begin
    FGapText := Value;
    AdjustBounds;
    Invalidate;
  end;
end;

procedure TInfoPanel.SetGapVertical(Value: Integer);
begin
  if (FGapVertical <> Value) and (Value >= 0) then
  begin
    FGapVertical := Value;
    AdjustBounds;
    Invalidate;
  end;
end;

procedure TInfoPanel.SetRadius(const Value: Integer);
begin
  if (FRadius <> Value) and (Value >= 0) then
  begin
    FRadius := Value;
    Invalidate;
  end;
end;

procedure TInfoPanel.SetShadow(Value: Boolean);
begin
  if FShadow <> Value then
  begin
    FShadow := Value;
    Invalidate;
  end;
end;

procedure TInfoPanel.SetShadowDepth(const Value: Integer);
begin
  if (FShadowDepth <> Value) and (Value >= 0) then
  begin
    FShadowDepth := Value;
    Invalidate;
  end;
end;

procedure TInfoPanel.SetTransparent(Value: Boolean);
begin
  if Transparent <> Value then
  begin
    if Value then
      ControlStyle := ControlStyle - [csOpaque]
    else
      ControlStyle := ControlStyle + [csOpaque];

    Invalidate;
  end;
end;

procedure TInfoPanel.SetWordWrap(Value: Boolean);
begin
  if FWordWrap <> Value then
  begin
    FWordWrap := Value;
    Invalidate;
  end;
end;

{$IFDEF VCL4}
procedure TInfoPanel.SetEnabled(Value: Boolean);
begin
  inherited;
  Invalidate;
end;
{$ENDIF}

end.
