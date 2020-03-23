unit QImport3WideStringCanvas;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF QI_UNICODE}

uses
  {$IFDEF VCL16}
    Vcl.Graphics,
    Vcl.Controls,
    Winapi.Windows;
  {$ELSE}
    Graphics,
    Controls,
    Windows;
  {$ENDIF}

type
   TEmsWideStringCanvas = class(TControlCanvas)
   public
     function TextExtentW(const Text: WideString): TSize;
     function TextHeightW(const Text: WideString): Integer;
     function TextHeight(const Text: WideString): Integer; overload;
     procedure TextOutW(X, Y: Integer; const Text: WideString);
    {$IFNDEF VCL15}
     procedure TextOut(X, Y: Integer; const Text: WideString);
        {$IFDEF VCL14} reintroduce; {$ELSE} overload; {$ENDIF}
    {$ENDIF}
     procedure TextRectW(Rect: TRect; X, Y: Integer; const Text: WideString);
     procedure TextRect(Rect: TRect; X, Y: Integer; const Text: WideString); overload;
     function TextWidthW(const Text: WideString): Integer;
     function TextWidth(const Text: WideString): Integer; overload;
   end;

{$ENDIF}

implementation

{$IFDEF QI_UNICODE}

{ TEmsWideStringCanvas }

function TEmsWideStringCanvas.TextExtentW(const Text: WideString): TSize;
begin
  RequiredState([csHandleValid, csFontValid]);
  Result.cX := 0;
  Result.cY := 0;
  {$IFDEF VCL16}Winapi.Windows{$ELSE}Windows{$ENDIF}.GetTextExtentPoint32W(Handle, PWideChar(Text), Length(Text), Result);
end;

function TEmsWideStringCanvas.TextHeightW(const Text: WideString): Integer;
begin
  Result := TextExtentW(Text).cY;
end;

function TEmsWideStringCanvas.TextHeight(const Text: WideString): Integer;
begin
  Result := TextExtentW(Text).cY;
end;

procedure TEmsWideStringCanvas.TextOutW(X, Y: Integer;
  const Text: WideString);
const
  MaxTextLength = 4094;
var
  TextLength: Integer;
  TempText,
  TextToDraw: WideString;
begin
  Changing;
  RequiredState([csHandleValid, csFontValid, csBrushValid]);
  
  TempText := Text;
  TextLength := Length(TempText);

  if CanvasOrientation = coRightToLeft then
  begin
    Inc(X, TextWidthW(TempText) + 1);
    while TextLength > MaxTextLength do
    begin
      TextToDraw := Copy(TempText, TextLength - MaxTextLength + 1, MaxTextLength);
      SetLength(TempText,  TextLength - MaxTextLength);
      TextLength := Length(TempText);
      {$IFDEF VCL16}Winapi.Windows{$ELSE}Windows{$ENDIF}.ExtTextOutW(Handle, X,
        Y, TextFlags, nil, PWideChar(TextToDraw), Length(TextToDraw), nil);
      Dec(X, TextWidthW(TextToDraw));
    end;
    TextToDraw := TempText;
    {$IFDEF VCL16}Winapi.Windows{$ELSE}Windows{$ENDIF}.ExtTextOutW(Handle, X,
      Y, TextFlags, nil, PWideChar(TextToDraw), Length(TextToDraw), nil);
    Dec(X, TextWidthW(TextToDraw));
    Inc(X, TextWidthW(Text) + 1);
  end
  else
  begin
    while TextLength > MaxTextLength do
    begin
      TextToDraw := Copy(TempText, 1, MaxTextLength);
      Delete(TempText,  1, MaxTextLength);
      TextLength := Length(TempText);
      {$IFDEF VCL16}Winapi.Windows{$ELSE}Windows{$ENDIF}.ExtTextOutW(Handle, X,
        Y, TextFlags, nil, PWideChar(TextToDraw), Length(TextToDraw), nil);
      Inc(X, TextWidthW(TextToDraw));
    end;
    TextToDraw := TempText;
    {$IFDEF VCL16}Winapi.Windows{$ELSE}Windows{$ENDIF}.ExtTextOutW(Handle, X,
      Y, TextFlags, nil, PWideChar(TextToDraw), Length(TextToDraw), nil);
    Inc(X, TextWidthW(TextToDraw));
  end;
  MoveTo(X, Y);
  Changed;
end;

{$IFNDEF VCL15}
procedure TEmsWideStringCanvas.TextOut(X, Y: Integer;
  const Text: WideString);
begin
  Changing;
  RequiredState([csHandleValid, csFontValid, csBrushValid]);
  if CanvasOrientation = coRightToLeft then
    Inc(X, TextWidth(Text) + 1);
  {$IFDEF VCL16}Winapi.Windows{$ELSE}Windows{$ENDIF}.ExtTextOutW(Handle, X, Y, TextFlags, nil, PWideChar(Text),
    Length(Text), nil);
  MoveTo(X + TextWidth(Text), Y);
  Changed;
end;
{$ENDIF}

procedure TEmsWideStringCanvas.TextRectW(Rect: TRect; X, Y: Integer;
  const Text: WideString);
var
  Options: Longint;
begin
  Changing;
  RequiredState([csHandleValid, csFontValid, csBrushValid]);
  Options := ETO_CLIPPED or TextFlags;
  if Brush.Style <> bsClear then
    Options := Options or ETO_OPAQUE;
  if ((TextFlags and ETO_RTLREADING) <> 0) and
      (CanvasOrientation = coRightToLeft) then
    Inc(X, TextWidthW(Text) + 1);
  {$IFDEF VCL16}Winapi.Windows{$ELSE}Windows{$ENDIF}.ExtTextOutW(Handle, X, Y, Options, @Rect, PWideChar(Text),
    Length(Text), nil);
  Changed;
end;

procedure TEmsWideStringCanvas.TextRect(Rect: TRect; X, Y: Integer;
  const Text: WideString);
var
  Options: Longint;
begin
  Changing;
  RequiredState([csHandleValid, csFontValid, csBrushValid]);
  Options := ETO_CLIPPED or TextFlags;
  if Brush.Style <> bsClear then
    Options := Options or ETO_OPAQUE;
  if ((TextFlags and ETO_RTLREADING) <> 0) and
      (CanvasOrientation = coRightToLeft) then
    Inc(X, TextWidth(Text) + 1);
  {$IFDEF VCL16}Winapi.Windows{$ELSE}Windows{$ENDIF}.ExtTextOutW(Handle, X, Y, Options, @Rect, PWideChar(Text),
    Length(Text), nil);
  Changed;
end;

function TEmsWideStringCanvas.TextWidthW(const Text: WideString): Integer;
begin
  Result := TextExtentW(Text).cX;
end;

function TEmsWideStringCanvas.TextWidth(const Text: WideString): Integer;
begin
  Result := TextExtentW(Text).cX;
end;

{$ENDIF}

end.
