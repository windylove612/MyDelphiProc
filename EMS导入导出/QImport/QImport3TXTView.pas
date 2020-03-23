unit QImport3TXTView;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    {$IFDEF QI_UNICODE}
      System.WideStrings,
      QImport3WideStringCanvas,
      QImport3CustomControl,
    {$ENDIF}
    Vcl.Controls,
    System.Classes,
    Vcl.StdCtrls,
    Winapi.Messages,
    Winapi.Windows,
  {$ELSE}
    {$IFDEF QI_UNICODE}
      {$IFDEF VCL10}
        WideStrings,
      {$ELSE}
        QImport3WideStrings,
      {$ENDIF}
      QImport3WideStringCanvas,
      QImport3CustomControl,
    {$ENDIF}
    Controls,
    Classes,
    StdCtrls,
    Messages,
    Windows,
  {$ENDIF}
  QImport3StrIDs,
  QImport3ASCII;

type
  TViewArrow = class(TCollectionItem)
  private
    FPosition: integer;
    procedure SetPosition(Value: integer);
  public
    constructor Create(Collection: TCollection); override;
    procedure Assign(Source: TPersistent); override;
    property Position: integer read FPosition write SetPosition;
  end;

  TQImport3TXTViewer = class;

  TViewArrows = class(TCollection)
  private
    FViewer: TQImport3TXTViewer;
    function GetItem(Index: integer): TViewArrow;
    procedure SetItem(Index: integer; Value: TViewArrow);
  public
    constructor Create(Viewer: TQImport3TXTViewer);
    function Add: TViewArrow;
{$IFNDEF VCL5}
    procedure Delete(Index: integer);
{$ENDIF}
//    procedure Sort;
//    function FindBetween(Start, Finish: integer; var Index: integer): boolean;
    function FindBetweenExcept(Start, Finish: integer; ExceptIndex: integer;
      var Index: integer): boolean;
    function FindLeftAndRight(Position: integer; var Left,
      Right: integer): boolean;
    function FindByPosition(Position: integer): boolean;

    property Items[Index: integer]: TViewArrow read GetItem
      write SetItem; default;
  end;

  TViewRulerAlign = (raTop, raBottom);
  TViewRuler = class(TGraphicControl)
  private
    FAlign: TViewRulerAlign;
    FOffset: integer;
    FStep: integer;

    procedure SetAlign(const Value: TViewRulerAlign);
    procedure SetOffset(Value: integer);
    procedure SetStep(Value: integer);
    procedure CMColorChanged(var Message: TMessage); message CM_COLORCHANGED;
  protected
    procedure Paint; override;

    property Offset: integer read FOffset write SetOffset;
    property Step: integer read FStep write SetStep;
  public
    constructor Create(AOwner: TComponent); override;
    property Align: TViewRulerAlign read FAlign write SetAlign;

    property Color;
  end;

  TViewSelection = class
  private
    FViewer: TQImport3TXTViewer;
    FVisibleRect: TRect;

    FLeftArrow: TViewArrow;
    FRightArrow: TViewArrow;

    FInverted: boolean;

    FOnChange: TNotifyEvent;

    function GetExists: boolean;
    procedure Update;
  public
    constructor Create(Viewer: TQImport3TXTViewer);
    procedure SetSelection(Left, Right: TViewArrow);

    property Viewer: TQImport3TXTViewer read FViewer;
    property LeftArrow: TViewArrow read FLeftArrow write FLeftArrow;
    property RightArrow: TViewArrow read FRightArrow write FRightArrow;
    property Exists: boolean read GetExists;
    property VisibleRect: TRect read FVisibleRect;

    property OnChange: TNotifyEvent read FOnChange write FOnChange;
  end;

  TDeleteArrowEvent = procedure(Sender: TObject; Position: integer) of object;
  TMoveArrowEvent = procedure(Sender: TObject; OldPos, NewPos: integer) of object;
  TIntersectArrowsEvent = procedure(Sender: TObject; Position: integer) of object;

  TQImport3TXTViewer = class({$IFDEF QI_UNICODE} TEmsCustomControl {$ELSE} TCustomControl {$ENDIF})
  private
    FImport: TQImport3ASCII;
    FCodePage: Integer;
    FScrollBars: TScrollStyle;

    FCharWidth: integer;
    FCharHeight: integer;
    FRealHeight: integer;
    FRealLeft: integer;
    FRealTop: integer;
    FRealWidth: integer;

    FTopRuler: TViewRuler;
    FBottomRuler: TViewRuler;

{$IFDEF QI_UNICODE}
    FLinesW: TWideStrings;
{$ELSE}
    FLines: TStrings;
{$ENDIF}

    FLineCount: integer;
    FMaxLineLength: integer;

    FArrows: TViewArrows;
    FActiveArrow: TViewArrow;
    FSelection: TViewSelection;

    FOnChangeSelection: TNotifyEvent;
    FOnDeleteArrow: TDeleteArrowEvent;
    FOnMoveArrow: TMoveArrowEvent;
    FOnIntersectArrows: TIntersectArrowsEvent;

    procedure SetScrollBars(const Value: TScrollStyle);
//    procedure SetLines(Value: TStrings);
    function CreateRuler(Align: TViewRulerAlign): TViewRuler;
    function GetFullWidth: integer;
    function GetFullHeight: integer;
    procedure WMHScroll(var Msg: TWMHScroll); message WM_HSCROLL;
    procedure WMVScroll(var Msg: TWMVScroll); message WM_VSCROLL;
    function GetNewArrowPosition(X: integer): integer;
    procedure ChangeSelection(Sender: TObject);
  protected
    procedure CreateParams(var Params: TCreateParams); override;
    procedure DrawRulers;
    procedure DrawLines;
    procedure DrawArrow(Index: integer);
    procedure DrawArrows;
    procedure Paint; override;
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState;
      X, Y: Integer); override;
    procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
    procedure MouseUp(Button: TMouseButton; Shift: TShiftState;
      X, Y: Integer); override;
  protected
    property ScrollBars: TScrollStyle read FScrollBars write SetScrollBars;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure LoadFromFile(const AFileName: string);
    procedure SetSelection(Pos, Size: integer);
    procedure GetSelection(var Pos, Size: integer);
    procedure AddArrow(Pos: integer);
    procedure DeleteArrows;

    property Arrows: TViewArrows read FArrows;
    property RealLeft: Integer read FRealLeft;
    property LineCount: integer read FLineCount write FLineCount;
    property CharWidth: Integer read FCharWidth;

    property OnChangeSelection: TNotifyEvent read FOnChangeSelection
      write FOnChangeSelection;
    property OnDeleteArrow: TDeleteArrowEvent read FOnDeleteArrow
      write FOnDeleteArrow;
    property OnMoveArrow: TMoveArrowEvent read FOnMoveArrow
      write FOnMoveArrow;
    property OnIntersectArrows: TIntersectArrowsEvent read FOnIntersectArrows
      write FOnIntersectArrows;
    property Import: TQImport3ASCII read FImport write FImport;
    property CodePage: Integer read FCodePage write FCodePage;
{$IFDEF QI_UNICODE}
    property Lines: TWideStrings read FLinesW;
{$ELSE}
    property Lines: TStrings read FLines;
{$ENDIF}
  end;

implementation

uses
  {$IFDEF VCL16}
    Vcl.Graphics,
    System.SysUtils,
  {$ELSE}
    Graphics,
    SysUtils,
  {$ENDIF}
  {$IFDEF QI_UNICODE}
    QImport3GpTextFile,
  {$ENDIF}
  QImport3,
  QImport3Common;

const
  sWidth    = 240;
  sHeight   = 120;
  sFontSize = 10;
  sFontName = 'Courier New';

{ TViewArrow }

constructor TViewArrow.Create(Collection: TCollection);
begin
  inherited;
  FPosition := -1;
end;

procedure TViewArrow.Assign(Source: TPersistent);
begin
  if Source is TViewArrow then
  begin
    Position := (Source as TViewArrow).Position;
    Exit;
  end;
  inherited;
end;

procedure TViewArrow.SetPosition(Value: integer);
begin
  if FPosition <> Value then
    FPosition := Value;
end;

{ TViewArrows }

constructor TViewArrows.Create(Viewer: TQImport3TXTViewer);
begin
  inherited Create(TViewArrow);
  FViewer := Viewer;
end;

function TViewArrows.Add: TViewArrow;
begin
  Result := (inherited Add) as TViewArrow;
end;

{$IFNDEF VCL5}
procedure TViewArrows.Delete(Index: integer);
begin
  TCollectionItem(Items[Index]).Free;
end;
{$ENDIF}

{procedure TViewArrows.Sort;

  procedure QuickSort(L, R: Integer);
  var
    i, j: Integer;
    P, T: TViewArrow;
  begin
    repeat
      i := L;
      j := R;
      P := Items[(L + R) shr 1];
      repeat
        while Items[i].Position < P.Position do
          Inc(i);
        while Items[j].Position > P.Position do
          Dec(j);
        if i <= j then
        begin
          T := Items[i];
          Items[i] := Items[j];
          Items[j] := T;
          Inc(i);
          Dec(j);
        end;
      until i > j;
      if L < j then
        QuickSort(L, j);
      L := i;
    until i >= R;
  end;

begin
  QuickSort(0, Count - 1);
end;}

{function TViewArrows.FindBetween(Start, Finish: integer;
  var Index: integer): boolean;
var
 L, H, I, C: Integer;
 j: integer;
begin
  Result := False;
  for j := Start to Finish do begin
    L := 0;
    H := Count - 1;
    while L <= H do begin
      I := (L + H) shr 1;
      if Items[i].Position < j then
        C := -1
      else if Items[i].Position > j then
        C := 1
      else C := 0;
      if C < 0 then L := I + 1
      else begin
        H := I - 1;
        if C = 0 then begin
          Result := True;
          L := I;
        end;
      end;
    end;
    Index := L;
    if Result then Break;
  end;
end;}

function TViewArrows.FindBetweenExcept(Start, Finish: integer;
  ExceptIndex: integer; var Index: integer): boolean;
var
  i: integer;
begin
  Result := False;
  for i := 0 to Count - 1 do
    if (i <> ExceptIndex) and (Items[i].Position >= Start) and
       (Items[i].Position <= Finish) then
    begin
      Result := True;
      Index := i;
      Break;
    end;
end;

function TViewArrows.FindLeftAndRight(Position: integer; var Left,
  Right: integer): boolean;
var
  i: integer;
begin
  Result := False;
  if Count < 2 then Exit;
  Left  := -1;
  Right := -1;
  for i := 0 to Count - 1 do
  begin
    if (Items[i].Position < Position) and
       ((Left = -1) or (Items[i].Position > Items[Left].Position)) then
      Left := i
    else if (Items[i].Position > Position) and
       ((Right = -1) or (Items[i].Position < Items[Right].Position)) then
      Right := i;
  end;
  Result := (Left > -1) and (Right > -1);
end;

function TViewArrows.FindByPosition(Position: integer): boolean;
var
  i: integer;
begin
  Result := False;
  for i := 0 to Count - 1 do
    if Items[i].Position = Position then
    begin
      Result := True;
      Break;
    end;
end;

function TViewArrows.GetItem(Index: integer): TViewArrow;
begin
  Result := inherited Items[Index] as TViewArrow;
end;

procedure TViewArrows.SetItem(Index: integer; Value: TViewArrow);
begin
  inherited SetItem(Index, Value);
end;

{ TViewRuler }

constructor TViewRuler.Create(AOwner: TComponent);
begin
  inherited;
  Height := 16;
  Width := 16;
  SetAlign(raTop);
  FStep := 10;
  ParentColor := False;
  Color := clBtnFace;
end;

procedure TViewRuler.Paint;

  procedure DrawVertTrack(X, Y1, Y2: integer);
  begin
    if X > 0 then
      with Canvas do
        case Align of
          raTop: begin
            Pen.Color := clBlack  ;
            Pen.Width := 1;
            MoveTo(X, Y1);
            LineTo(X, Y2);
            Pen.Color := clWhite;
            MoveTo(X + 1, Y1);
            LineTo(X + 1, Y2);
          end;
          raBottom: begin
            Pen.Color := clBlack;
            Pen.Width := 1;
            MoveTo(X, Y1);
            LineTo(X, Y2 - 1);
            Pen.Color := clWhite;
            MoveTo(X + 1, Y1);
            LineTo(X + 1, Y2 - 1);
          end
        end;
  end;

  procedure DrawBorder;
  begin
    with Canvas do
    begin
      case Align of
        raTop: begin
          Pen.Color := clBlack;
          Pen.Width := 1;
          MoveTo(0, Height - 1);
          LineTo(Width - 1, Height - 1);

          Pen.Color := clWhite;
          Pen.Width := 1;
          MoveTo(0, 0);
          LineTo(Width - 1, 0);
          MoveTo(0, 0);
          LineTo(0, Height);

          Pen.Color := clBlack;
          Pen.Width := 1;
          MoveTo(Width - 1, 0);
          LineTo(Width - 1, Height - 1);
        end;
        raBottom: begin
          Pen.Color := clBlack;
          Pen.Width := 1;
          MoveTo(0, Height - 1);
          LineTo(Width - 1, Height - 1);


          Pen.Color := clBlack;
          Pen.Width := 1;
          MoveTo(0, 0);
          LineTo(Width - 1, 0);
          Pen.Color := clWhite;
          Pen.Width := 1;
          MoveTo(0, 1);
          LineTo(Width - 1, 1);

          Pen.Color := clWhite;
          Pen.Width := 1;
          MoveTo(0, 1);
          LineTo(0, Height - 1);

          Pen.Color := clBlack;
          Pen.Width := 1;
          MoveTo(Width - 1, 1);
          LineTo(Width - 1, Height);
        end;
      end;
    end;
  end;

var
  i, ASize: integer;
  HDivisor: byte;
begin
  inherited Paint;

  Canvas.Brush.Color := clBtnFace;
  Canvas.FillRect(Rect(0, 0, Width, Height));

  DrawBorder;

  if FStep > 0
    then ASize := (Width * 2 - 1) div FStep
    else ASize := 0;

  for i := 1 to ASize do begin
    if (i = FOffset) or ((i + FOffset - 1) mod 5 = 0)
      then HDivisor := 1
      else HDivisor := 2;
    case FAlign of
      raBottom : DrawVertTrack(i * FStep, 2,
                              (Height - 1) div HDivisor);
      raTop    : DrawVertTrack(i * FStep, Height - 3,
                               Height - (Height - 1) div HDivisor);
    end;
  end;
end;

procedure TViewRuler.SetAlign(const Value: TViewRulerAlign);
begin
  if FAlign <> Value then
    FAlign := Value;
  case FAlign of
    raBottom : inherited Align := alBottom;
    raTop    : inherited Align := alTop;
  end;
end;

procedure TViewRuler.SetOffset(Value: integer);
begin
  if FOffset <> Value then
  begin
    FOffset := Value;
    Invalidate;
  end;
end;

procedure TViewRuler.SetStep(Value: integer);
begin
  if FStep <> Value then
  begin
    FStep := Value;
    Invalidate;
  end;
end;

procedure TViewRuler.CMColorChanged(var Message: TMessage);
begin
  Canvas.Brush.Color := Color;
  inherited;
end;

{ TViewSelection }

constructor TViewSelection.Create(Viewer: TQImport3TXTViewer);
begin
  inherited Create;
  FViewer := Viewer;
  FLeftArrow := nil;
  FRightArrow := nil;
  FVisibleRect := Rect(0, 0, 0, 0);
  FInverted := False;
end;

procedure TViewSelection.Update;
var
  R: TRect;
begin
  if Assigned(FLeftArrow) and Assigned(FRightArrow) then
  begin
    if FRightArrow.Position > 0 then
    begin
     if FRightArrow.Position < FViewer.ClientWidth
       then R.Right := FRightArrow.Position
       else R.Right := FViewer.ClientWidth;
    end
    else R.Right := 0;

    if FLeftArrow.Position > 0 then
    begin
     if FLeftArrow.Position < FViewer.ClientWidth
       then R.Left := FLeftArrow.Position
       else R.Left := FViewer.ClientWidth;
    end
    else R.Left := 0;

    R.Top := FViewer.FTopRuler.Height;
    R.Bottom := FViewer.ClientHeight - FViewer.FBottomRuler.Height;

    if not Assigned(FViewer.FActiveArrow) then
    begin
      if (R.Right > FVisibleRect.Right) and
         (R.Left = FVisibleRect.Left) then
      begin
        InvertRect(FViewer.Canvas.Handle,
          Rect(R.Left, R.Top, R.Right - FVisibleRect.Right, R.Bottom));
        FInverted := True;
      end
      else if (R.Left < FVisibleRect.Left) and
          (R.Right = FVisibleRect.Right) then
      begin
        InvertRect(FViewer.Canvas.Handle,
        Rect(R.Right - (R.Left - FVisibleRect.Left), R.Top, R.Right, R.Bottom));
        FInverted := True;
      end;
    end
    else begin
      if (R.Right > FVisibleRect.Right) and
         (R.Left = FVisibleRect.Left) then
      begin
        InvertRect(FViewer.Canvas.Handle,
          Rect(FVisibleRect.Right, R.Top, R.Right, R.Bottom));
        FInverted := True;
      end
      else if (R.Right < FVisibleRect.Right) and
          (R.Left = FVisibleRect.Left) then
      begin
        InvertRect(FViewer.Canvas.Handle,
          Rect(R.Right, R.Top, FVisibleRect.Right, R.Bottom));
        FInverted := True;
      end
      else if (R.Left < FVisibleRect.Left) and
          (R.Right = FVisibleRect.Right) then
      begin
        InvertRect(FViewer.Canvas.Handle,
          Rect(R.Left, R.Top, FVisibleRect.Left, R.Bottom));
        FInverted := True;
      end
      else if (R.Left > FVisibleRect.Left) and
          (R.Right = FVisibleRect.Right) then
      begin
        InvertRect(FViewer.Canvas.Handle,
          Rect(FVisibleRect.Left, R.Top, R.Left, R.Bottom));
        FInverted := True;
      end;
    end;

    FVisibleRect := R;
    if not FInverted then
    begin
      InvertRect(FViewer.Canvas.Handle, FVisibleRect);
      FInverted := True;
    end;
  end
  else begin
    InvertRect(FViewer.Canvas.Handle, FVisibleRect);
    FVisibleRect := Rect(0, 0, 0, 0);
    FInverted := False;
  end;
  if Assigned(FOnChange) then FOnChange(Self);
end;

function TViewSelection.GetExists: boolean;
begin
  Result := Assigned(FLeftArrow) and Assigned(FRightArrow);
end;

procedure TViewSelection.SetSelection(Left, Right: TViewArrow);
begin
  if (FLeftArrow <> Left) or (FRightArrow <> Right) then
  begin
    FLeftArrow := Left;
    FRightArrow := Right;
    Update;
  end;
end;

{ TQImport3TXTViewer }

constructor TQImport3TXTViewer.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  Height := sHeight;
  Width := sWidth;
  FScrollBars := ssNone;

  Color := clWhite;
  with Canvas do
  begin
    Brush.Color := Color;
    Pen.Color := clBlack;
    Canvas.Pen.Mode := pmNot;
    Font.Name := sFontName;
    Font.Size := sFontSize;
  end;

  FCharWidth := 0;
  FCharHeight := 0;
  FRealHeight := Height;
  FRealLeft := 0;
  FRealTop := 0;
  FRealWidth := Width;

  FTopRuler := CreateRuler(raTop);
  FBottomRuler := CreateRuler(raBottom);

  FLineCount := 100;
  FMaxLineLength := 0;
  FCodePage := -1;
{$IFDEF QI_UNICODE}
  FLinesW := TWideStringList.Create;
{$ELSE}
  FLines := TStringList.Create;
{$ENDIF}

  FArrows := TViewArrows.Create(Self);
  FActiveArrow := nil;

  FSelection := TViewSelection.Create(Self);
  FSelection.OnChange := ChangeSelection;
end;

destructor TQImport3TXTViewer.Destroy;
begin
  FSelection.Free;
  FArrows.Free;
{$IFDEF QI_UNICODE}
  FLinesW.Free;
{$ELSE}
  FLines.Free;
{$ENDIF}
  inherited;
end;

procedure TQImport3TXTViewer.LoadFromFile(const AFileName: string);
var
{$IFDEF QI_UNICODE}
  tf: TGpTextFile;
  wstr: WideString;
{$ELSE}
  F: TextFile;
  str: AnsiString;
{$ENDIF}
  i, l: integer;
begin
  if AFileName = EmptyStr then
    raise Exception.Create(QImportLoadStr(QIE_NoFileName));
  if not FileExists(AFileName) then
    raise Exception.CreateFmt(QImportLoadStr(QIE_FileNotExists), [AFileName]);

{$IFDEF QI_UNICODE}
  tf := TGpTextFile.Create(AFileName);
  tf.TryReadUpCRLF := True;
  try
    tf.Reset;
    if CodePage <> -1 then
      tf.Codepage := Codepage
    else
      Codepage := tf.Codepage;
    FLinesW.Clear;
    FMaxLineLength := 0;
    i := 0;
    while not tf.Eof do
    begin
      if (FLineCount > 0) and (i >= LineCount) then Break;
      wstr := tf.Readln;
      ReplaceTabs(wstr);
      FLinesW.Add(wstr);
      l := Length(wstr);
      if l > FMaxLineLength then
        FMaxLineLength := l;
      Inc(i);
    end;
  finally
    tf.Free;
  end;
{$ELSE}
  AssignFile(F, AFileName);
  Reset(F);
  try
    FLines.Clear;
    FMaxLineLength := 0;
    i := 0;
    while not Eof(F) do
    begin
      if (FLineCount > 0) and (i >= LineCount) then Break;
      Readln(F, str);
      ReplaceTabs(str);
      FLines.Add(str);
      l := Length(str);
      if l > FMaxLineLength then
        FMaxLineLength := l;
      Inc(i);
    end;
  finally
    CloseFile(F);
  end;
{$ENDIF}

  Invalidate;

  try
    if FCharHeight = 0 then
    begin
{$IFDEF QI_UNICODE}
      FCharHeight := Canvas.TextHeightW('A');
{$ELSE}
      FCharHeight := Canvas.TextHeight('A');
{$ENDIF}
    end;
         
    if FCharWidth = 0 then
    begin
{$IFDEF QI_UNICODE}
      FCharWidth := Canvas.TextWidthW('A');
{$ELSE}
      FCharWidth := Canvas.TextWidth('A');
{$ENDIF}
    end;
  except
  end;
end;

procedure TQImport3TXTViewer.SetSelection(Pos, Size: integer);
var
  L, R, Al, Ar: integer;
begin
  FSelection.SetSelection(nil, nil);
  L := (Pos + 1) * FCharWidth + FRealLeft;
  R := (Pos + Size + 1) * FCharWidth + FRealLeft;

  if FArrows.FindBetweenExcept(L, L, -1, Al) and
     FArrows.FindBetweenExcept(R, R, -1, Ar) then
  begin
    FSelection.SetSelection(FArrows[Al], FArrows[Ar]);
  end;
end;

procedure TQImport3TXTViewer.GetSelection(var Pos, Size: integer);
begin
  if FSelection.Exists then
  begin
    Pos := ((FSelection.LeftArrow.Position - FRealLeft) div FCharWidth) - 1;
    Size := ((FSelection.RightArrow.Position - FRealLeft) div FCharWidth) - 1 - Pos;
  end
  else begin
    Pos := -1;
    Size := -1;
  end;
end;

procedure TQImport3TXTViewer.AddArrow(Pos: integer);
var
  Position: integer;
begin
  Position := (Pos + 1) * FCharWidth + FRealLeft;
  if not FArrows.FindByPosition(Position) then
    FArrows.Add.Position := Position;
end;

procedure TQImport3TXTViewer.DeleteArrows;
begin
  while FArrows.Count > 0 do
    FArrows[0].Free;
end;

procedure TQImport3TXTViewer.CreateParams(var Params: TCreateParams);
const
  ScrollBar: array[TScrollStyle] of DWORD = (0, WS_HSCROLL, WS_VSCROLL,
    WS_HSCROLL or WS_VSCROLL);
begin
  inherited CreateParams(Params);
  with Params do
  begin
    Style := Style or ScrollBar[FScrollBars];
    if NewStyleControls and Ctl3D then
    begin
      Style := Style and (not WS_BORDER);
      ExStyle := ExStyle or WS_EX_CLIENTEDGE;
    end
    else
      Style := Style or WS_BORDER;
    WindowClass.Style := WindowClass.Style and not (CS_HREDRAW or CS_VREDRAW);
  end;
end;

procedure TQImport3TXTViewer.DrawRulers;
begin
  if FBottomRuler.Step = 0 then
    FBottomRuler.Step := FCharWidth;
  if FTopRuler.Step = 0 then
    FTopRuler.Step := FCharWidth;
end;

procedure TQImport3TXTViewer.DrawLines;
var
  i, t, cnt: integer;
  ScrollInfo: TScrollInfo;
begin
{$IFDEF QI_UNICODE}
  cnt := FLinesW.Count;
{$ELSE}
  cnt := FLines.Count;
{$ENDIF}

  for i := 0 to cnt - 1 do
  begin
    t := FRealTop + FTopRuler.Height + 3 + i * (FCharHeight + 1);
    if (t > 0) and (t < ClientHeight - FBottomRuler.Height - FTopRuler.Height) then
    begin
{$IFDEF QI_UNICODE}
      Canvas.TextOutW(FRealLeft + FCharWidth, t, FLinesW[i] + ' ');
{$ELSE}
      Canvas.TextOut(FRealLeft + FCharWidth, t, FLines[i] + ' ');
{$ENDIF}
    end;
  end;

  if FSelection.Exists then
    InvertRect(Canvas.Handle, FSelection.VisibleRect);

  if GetFullWidth > ClientWidth then
  begin
    if ScrollBars = ssVertical then
      ScrollBars := ssBoth
    else if ScrollBars = ssNone then
      ScrollBars := ssHorizontal;

    ScrollInfo.cbSize := SizeOf(ScrollInfo);
    ScrollInfo.fMask := SIF_ALL;
    GetScrollInfo(Self.Handle, SB_HORZ, ScrollInfo);
    ScrollInfo.nMin := 0;
    if FCharWidth = 0
      then ScrollInfo.nMax := 0
      else ScrollInfo.nMax := FMaxLineLength - 1;
    SetScrollInfo(Self.Handle, SB_HORZ, ScrollInfo, True);
  end;
  if GetFullHeight > ClientHeight then
  begin
    if ScrollBars = ssHorizontal then
      ScrollBars := ssBoth
    else if ScrollBars = ssNone then
      ScrollBars := ssVertical;

    ScrollInfo.cbSize := SizeOf(ScrollInfo);
    ScrollInfo.fMask := SIF_ALL;
    GetScrollInfo(Self.Handle, SB_VERT, ScrollInfo);
    ScrollInfo.nMin := 0;
    if FCharHeight = 0
      then ScrollInfo.nMax := 0
      else ScrollInfo.nMax := GetFullHeight div FCharHeight - 1;
    SetScrollInfo(Self.Handle, SB_VERT, ScrollInfo, True);
  end;
end;

procedure TQImport3TXTViewer.DrawArrow(Index: integer);
var
  ATop, AHeight: integer;
begin
  ATop := 16;
  AHeight := ClientHeight - ATop;
  Canvas.MoveTo(FArrows[Index].Position, ATop);
  Canvas.LineTo(FArrows[Index].Position, AHeight);
end;

procedure TQImport3TXTViewer.DrawArrows;
var
  i, l: integer;
begin
  for i := 0 to FArrows.Count - 1 do
  begin
    l := GetNewArrowPosition(FArrows[i].Position);
    if l > 0 then
      FArrows[i].Position := l;
    DrawArrow(i);
  end;
end;

procedure TQImport3TXTViewer.Paint;
begin
  inherited;
  if Assigned(Parent) then
  begin
    if FCharHeight = 0 then
    begin
{$IFDEF QI_UNICODE}
      FCharHeight := Canvas.TextHeightW('A');
{$ELSE}
      FCharHeight := Canvas.TextHeight('A');
{$ENDIF}
    end;

    if FCharWidth = 0 then
    begin
{$IFDEF QI_UNICODE}
      FCharWidth := Canvas.TextWidthW('A');
{$ELSE}
      FCharWidth := Canvas.TextWidth('A');
{$ENDIF}
    end;

    FRealHeight := GetFullHeight;
    FRealWidth := GetFullWidth;
    DrawRulers;
    DrawLines;
    if not Assigned(FActiveArrow) then DrawArrows;
  end;
end;

procedure TQImport3TXTViewer.MouseDown(Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  Index, L, R, L1, R1: integer;
  Arrow: TViewArrow;
begin
  inherited;
  if not (ssLeft in Shift) then Exit;
  if ssDouble in Shift then
  begin
    if FArrows.FindBetweenExcept(X - 3, X + 3, -1, Index) then
    begin
      if FSelection.Exists and
         ((FArrows[Index] = FSelection.LeftArrow) or
          (FArrows[Index] = FSelection.RightArrow)) then
        FSelection.SetSelection(nil, nil);
      DrawArrow(Index);
      if Assigned(FOnDeleteArrow) then
        FOnDeleteArrow(Self, (FArrows[Index].Position - FRealLeft) div FCharWidth - 1);
      FArrows.Delete(Index);
      Cursor := crDefault;
    end
    else begin
      L := GetNewArrowPosition(X);
      if ((L - FRealLeft) div FCharWidth) < 1 then Exit;
      Arrow := FArrows.Add;
      Arrow.Position := GetNewArrowPosition(X);
      DrawArrow(Arrow.Index);
      if FSelection.Exists and
         ((Arrow.Position > FSelection.LeftArrow.Position) and
          (Arrow.Position < FSelection.RightArrow.Position)) then
        FSelection.SetSelection(nil, nil);
      //FArrows.Sort;
    end;
  end
  else begin
    if FArrows.FindBetweenExcept(X - 3, X + 3, -1, Index) then
      FActiveArrow := FArrows[Index]
    else begin
      if FArrows.FindLeftAndRight(X, L, R) then
      begin
        L1 := -1;
        R1 := -1;
        if FSelection.Exists then
        begin
          L1 := FSelection.LeftArrow.Index;
          R1 := FSelection.RightArrow.Index;
          FSelection.SetSelection(nil, nil);
        end;
        if (L <> L1) and (R <> R1) then
          FSelection.SetSelection(FArrows[L], FArrows[R]);
      end;
    end;
  end
end;

procedure TQImport3TXTViewer.MouseMove(Shift: TShiftState; X, Y: Integer);
var
  OP, NP, Index: integer;
begin
  inherited;
  if Shift = [] then
  begin
    if FArrows.FindBetweenExcept(X - 3, X + 3, -1, Index)
      then Cursor := crHSplit
      else Cursor := crDefault;
  end
  else if (ssLeft in Shift) and Assigned(FActiveArrow) then
  begin
    if X - FRealLeft < FCharWidth then // ab
      Exit;
    OP := FActiveArrow.Position;
    if not FArrows.FindBetweenExcept(OP, Op, FActiveArrow.Index, Index) then
      DrawArrow(FActiveArrow.Index);
    NP := GetNewArrowPosition(X);
    FActiveArrow.Position := NP;
    if Assigned(FOnMoveArrow) then
      FOnMoveArrow(Self, (OP - FRealLeft) div FCharWidth - 1,
                         (NP - FRealLeft) div FCharWidth - 1);
    if not FArrows.FindBetweenExcept(NP, NP, FActiveArrow.Index, Index) then
    begin
      DrawArrow(FActiveArrow.Index);
      if (FActiveArrow = FSelection.LeftArrow) or
         (FActiveArrow = FSelection.RightArrow) then
      begin
        FSelection.Update;
      end;
    end
    else begin
      FSelection.SetSelection(nil, nil);
      if Assigned(FOnIntersectArrows) then
        FOnIntersectArrows(Self, (NP - FRealLeft) div FCharWidth - 1);
    end;
  end;
end;

procedure TQImport3TXTViewer.MouseUp(Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  L, Index: integer;
begin
  inherited;
  if Assigned(FActiveArrow) then
  begin
    if X - FRealLeft >= FCharWidth then // ab
    begin
      L := GetNewArrowPosition(X);
      if FArrows.FindBetweenExcept(L, L, FActiveArrow.Index, Index) then
        FArrows.Delete(FActiveArrow.Index)
      else if ((L - FRealLeft) div FCharWidth) < 1 then
      begin
        DrawArrow(FActiveArrow.Index);
        FArrows.Delete(FActiveArrow.Index);
      end;
      //  else DrawArrow(FActiveArrow.Index);
    end;
    FActiveArrow := nil;
    //FArrows.Sort;
  end;
end;

procedure TQImport3TXTViewer.SetScrollBars(const Value: TScrollStyle);
begin
  if FScrollBars <> Value then
  begin
    FScrollBars := Value;
    RecreateWnd;
  end;
end;

function TQImport3TXTViewer.CreateRuler(Align: TViewRulerAlign): TViewRuler;
begin
  Result := TViewRuler.Create(Self);
  Result.Parent := Self;
  Result.Align := Align;
  Result.Step := FCharWidth;
end;

function TQImport3TXTViewer.GetFullWidth: integer;
begin
  Result := FCharWidth * FMaxLineLength;
end;

function TQImport3TXTViewer.GetFullHeight: integer;
begin
  Result := FCharHeight * {$IFDEF QI_UNICODE}FLinesW{$ELSE}FLines{$ENDIF}.Count;
//  Result := ClientHeight - FRealTop - FBottomRuler.Height;
end;

procedure TQImport3TXTViewer.WMHScroll(var Msg: TWMHScroll);

  procedure HorzScroll(Step: integer);
  begin
    FBottomRuler.Offset := FBottomRuler.Offset + (FCharWidth + 1) * Step;
    FTopRuler.Offset := FTopRuler.Offset + (FCharWidth + 1) * Step;
  end;

var
  OldScroll, NewScroll: TScrollInfo;
  i, Step: integer;
begin
  OldScroll.cbSize := SizeOf(OldScroll);
  OldScroll.fMask := SIF_ALL;
  GetScrollInfo(Self.Handle, SB_HORZ, OldScroll);
  OldScroll.nMin := 0;
  if FCharWidth = 0
    then OldScroll.nMax := 0
    else OldScroll.nMax := FMaxLineLength - 1;
  NewScroll := OldScroll;
  case Msg.ScrollCode of
    SB_LINELEFT,
    SB_PAGELEFT:
      if NewScroll.nPos > NewScroll.nMin then
      begin
        if Msg.ScrollCode = SB_LINELEFT
          then Step := 1
          else Step := 10;
        if NewScroll.nPos <= NewScroll.nMin + Step then
          Step := NewScroll.nPos - NewScroll.nMin;
        Dec(NewScroll.nPos, Step);
        Inc(FRealLeft, FCharWidth * Step);
        ScrollBy(FCharWidth * Step, 0);
        HorzScroll(Step);
        for i := 0 to FArrows.Count - 1 do
          FArrows[i].Position := FArrows[i].Position + FCharWidth * Step;
      end;
    SB_LINERIGHT,
    SB_PAGERIGHT:
      if NewScroll.nPos < NewScroll.nMax then
      begin
        if Msg.ScrollCode = SB_LINERIGHT
          then Step := 1
          else Step := 10;
        if NewScroll.nPos >= NewScroll.nMax - Step then
          Step := NewScroll.nMax - NewScroll.nPos;
        Inc(NewScroll.nPos, Step);
        Dec(FRealLeft, FCharWidth * Step);
        ScrollBy(-FCharWidth * Step, 0);
        HorzScroll(-Step);
        for i := 0 to FArrows.Count - 1 do
          FArrows[i].Position := FArrows[i].Position - FCharWidth * Step;
      end;
    SB_THUMBTRACK:
      if NewScroll.nPos <> NewScroll.nTrackPos then
      begin
        FRealLeft := FRealLeft -
          (NewScroll.nTrackPos - NewScroll.nPos) * FCharWidth;
        ScrollBy((NewScroll.nPos - NewScroll.nTrackPos) * FCharWidth, 0);
        HorzScroll(NewScroll.nPos - NewScroll.nTrackPos);
        for i := 0 to FArrows.Count - 1 do
          FArrows[i].Position := FArrows[i].Position -
            (NewScroll.nTrackPos - NewScroll.nPos) * FCharWidth;
        NewScroll.nPos := NewScroll.nTrackPos;
      end;
    else Exit;
  end;
  if @OldScroll <> @NewScroll then
    SetScrollInfo(Self.Handle, SB_HORZ, NewScroll, True);
  FSelection.Update;
end;

procedure TQImport3TXTViewer.WMVScroll(var Msg: TWMVScroll);
var
  OldScroll, NewScroll: TScrollInfo;
  Step: integer;
begin
  OldScroll.cbSize := SizeOf(OldScroll);
  OldScroll.fMask := SIF_ALL;
  GetScrollInfo(Self.Handle, SB_VERT, OldScroll);
  OldScroll.nMin := 0;
  if FCharWidth = 0 then
    OldScroll.nMax := 0
  else
    OldScroll.nMax := {$IFDEF QI_UNICODE}FLinesW{$ELSE}FLines{$ENDIF}.Count - 1;
  NewScroll := OldScroll;
  case Msg.ScrollCode of
    SB_LINEUP,
    SB_PAGEUP:
      if NewScroll.nPos > NewScroll.nMin then
      begin
        if Msg.ScrollCode = SB_LINEUP
          then Step := 1
          else Step := 10;
        if NewScroll.nPos <= NewScroll.nMin + Step then
          Step := NewScroll.nPos - NewScroll.nMin;
        Dec(NewScroll.nPos, Step);
        Inc(FRealTop, FCharHeight * Step);
        ScrollBy(0, FCharHeight * Step);
      end;
    SB_LINEDOWN,
    SB_PAGEDOWN:
      if NewScroll.nPos < NewScroll.nMax then
      begin
        if Msg.ScrollCode = SB_LINEDOWN
          then Step := 1
          else Step := 10;
        if NewScroll.nPos >= NewScroll.nMax - Step then
          Step := NewScroll.nMax - NewScroll.nPos;
        Inc(NewScroll.nPos, Step);
        Dec(FRealTop, FCharHeight * Step);
        ScrollBy(0, -FCharHeight * Step);
      end;
    SB_THUMBTRACK:
      if NewScroll.nPos <> NewScroll.nTrackPos then
      begin
        FRealTop := FRealTop -
          (NewScroll.nTrackPos - NewScroll.nPos) * FCharHeight;
        ScrollBy(0, (NewScroll.nPos - NewScroll.nTrackPos) * FCharHeight);
        NewScroll.nPos := NewScroll.nTrackPos;
      end;
    else Exit;
  end;
  if @OldScroll <> @NewScroll then
    SetScrollInfo(Self.Handle, SB_VERT, NewScroll, True);
end;

function TQImport3TXTViewer.GetNewArrowPosition(X: integer): integer;
begin
  Result := 0;
  if (FCharWidth > 0) and (X > 0) then
  begin
    if (X mod FCharWidth) > (FCharWidth / 2) then
      Result := ((X div FCharWidth) + 1) * FCharWidth
    else Result := (X div FCharWidth) * FCharWidth;
  end
end;

procedure TQImport3TXTViewer.ChangeSelection(Sender: TObject);
begin
  if Assigned(FOnChangeSelection) then FOnChangeSelection(Self);
end;

end.
