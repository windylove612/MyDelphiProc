unit QImport3WideStrUtils;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    System.SysUtils;
  {$ELSE}
    Windows,
    SysUtils;
  {$ENDIF}



{$IFDEF QI_UNICODE}
function QICompareStr(const S1, S2: WideString): Integer;
function QICompareText(const S1, S2: WideString): Integer;
function QIUpperCase(const S: WideString): WideString;
function QILowerCase(const S: WideString): WideString;
function QIPos(const Substr, S: WideString): Integer;
procedure QIInsert(const Substr: WideString; var S: WideString; const Index: Integer);
procedure QIDelete(var S: WideString; const Index, Count: Integer);
function QIFormat(const Format: WideString; const Args: array of const): WideString;
function QIStringReplace(const S, OldPattern, NewPattern: WideString;
  Flags: TReplaceFlags): WideString;
function QIQuotedStr(const S: WideString; Quote: WideChar): WideString;
{$ELSE}
function QICompareStr(const S1, S2: string): Integer;
function QICompareText(const S1, S2: string): Integer;
function QIUpperCase(const S: string): string;
function QILowerCase(const S: string): string;
function QIPos(const Substr, S: string): Integer;
procedure QIInsert(const Substr: string; var S: string; const Index: Integer);
procedure QIDelete(var S: string; const Index, Count: Integer);
function QIFormat(const Format: string; const Args: array of const): string;
function QIStringReplace(const S, OldPattern, NewPattern: string;
  Flags: TReplaceFlags): string;
function QIQuotedStr(const S: string; Quote: Char): string;
{$ENDIF}

{$IFDEF QI_UNICODE}
{$IFNDEF VCL9}
procedure WideFmtStr(var Result: WideString; const Format: WideString;
  const Args: array of const);
function WideFormat(const Format: WideString; const Args: array of const): WideString;
function WideStringReplace(const S, OldPattern, NewPattern: Widestring;
  Flags: TReplaceFlags): Widestring;
function WideReplaceStr(const AText, AFromText, AToText: WideString): WideString;
function WideReplaceText(const AText, AFromText, AToText: WideString): WideString;
function WStrAlloc(Size: Cardinal): PWideChar;
function WStrBufSize(const Str: PWideChar): Cardinal;
{$ENDIF}
{$IFNDEF VCL10}
function WStrMove(Dest: PWideChar; const Source: PWideChar; Count: Cardinal): PWideChar;
{$ENDIF}
{$IFNDEF VCL9}
function WStrNew(const Str: PWideChar): PWideChar;
procedure WStrDispose(Str: PWideChar);
{$ENDIF}
{$IFNDEF VCL9}
function WStrLen(Str: PWideChar): Cardinal;
function WStrEnd(Str: PWideChar): PWideChar;
{$ENDIF}
{$IFNDEF VCL10}
function WStrCat(Dest: PWideChar; const Source: PWideChar): PWideChar;
{$ENDIF}
{$IFNDEF VCL9}
function WStrCopy(Dest, Source: PWideChar): PWideChar;
function WStrLCopy(Dest, Source: PWideChar; MaxLen: Cardinal): PWideChar;
function WStrPCopy(Dest: PWideChar; const Source: AnsiString): PWideChar;
function WStrPLCopy(Dest: PWideChar; const Source: AnsiString; MaxLen: Cardinal): PWideChar;
{$ENDIF}
{$IFNDEF VCL10}
function WStrScan(const Str: PWideChar; Chr: WideChar): PWideChar;
function WStrComp(Str1, Str2: PWideChar): Integer;
function WStrPos(Str, SubStr: PWideChar): PWideChar;
{$ENDIF}
function WStrECopy(Dest, Source: PWideChar): PWideChar;
function WStrLComp(Str1, Str2: PWideChar; MaxLen: Cardinal): Integer;
function WStrLIComp(Str1, Str2: PWideChar; MaxLen: Cardinal): Integer;
function WStrIComp(Str1, Str2: PWideChar): Integer;
function WStrLower(Str: PWideChar): PWideChar;
function WStrUpper(Str: PWideChar): PWideChar;
function WStrRScan(const Str: PWideChar; Chr: WideChar): PWideChar;
function WStrLCat(Dest: PWideChar; const Source: PWideChar; MaxLen: Cardinal): PWideChar;
function WStrPas(const Str: PWideChar): WideString;

{$IFNDEF VCL10}
function WideLastChar(const S: WideString): PWideChar;
function WideQuotedStr(const S: WideString; Quote: WideChar): WideString;
{$ENDIF}
{$IFNDEF VCL9}
function WideExtractQuotedStr(var Src: PWideChar; Quote: WideChar): Widestring;
{$ENDIF}
{$IFNDEF VCL10}
function WideDequotedStr(const S: WideString; AQuote: WideChar): WideString;
{$ENDIF}

{$ENDIF}

implementation

uses
  {$IFDEF VCL16}
    System.WideStrUtils,
    Winapi.Windows,
    System.Math;
  {$ELSE}
    {$IFDEF VCL9}
      WideStrUtils,
    {$ENDIF}
    Math;
  {$ENDIF}


{$IFDEF QI_UNICODE}

{$IFDEF VER130}
type
  _EOSError = class(Exception)
  public
    ErrorCode: DWORD;
  end;

resourcestring
  _SOSError = 'System Error.  Code: %d.'+#10+'%s';
  _SUnkOSError = 'A call to an OS function failed';

function _DumbItDownFor95(const S1, S2: WideString; CmpFlags: Integer): Integer;
var
  a1, a2: AnsiString;
begin
  a1 := s1;
  a2 := s2;
  Result := CompareStringA(LOCALE_USER_DEFAULT, CmpFlags, PChar(a1), Length(a1),
    PChar(a2), Length(a2)) - 2;
end;

procedure _RaiseLastOSError;
var
  LastError: Integer;
  Error: _EOSError;
begin
  LastError := GetLastError;
  if LastError <> 0 then
    Error := _EOSError.CreateResFmt(@_SOSError, [LastError,
      SysErrorMessage(LastError)])
  else
    Error := _EOSError.CreateRes(@_SUnkOSError);
  Error.ErrorCode := LastError;
  raise Error;
end;

function _WideCompareStr(const S1, S2: WideString): Integer;
begin
  SetLastError(0);
  Result := CompareStringW(LOCALE_USER_DEFAULT, 0, PWideChar(S1), Length(S1),
    PWideChar(S2), Length(S2)) - 2;
  case GetLastError of
    0: ;
    ERROR_CALL_NOT_IMPLEMENTED: Result := _DumbItDownFor95(S1, S2, 0);
  else
    _RaiseLastOSError;
  end;
end;

function _WideCompareText(const S1, S2: WideString): Integer;
begin
  SetLastError(0);
  Result := CompareStringW(LOCALE_USER_DEFAULT, NORM_IGNORECASE, PWideChar(S1),
    Length(S1), PWideChar(S2), Length(S2)) - 2;
  case GetLastError of
    0: ;
    ERROR_CALL_NOT_IMPLEMENTED: Result := _DumbItDownFor95(S1, S2, NORM_IGNORECASE);
  else
    _RaiseLastOSError;
  end;
end;

function _WideUpperCase(const S: WideString): WideString;
var
  Len: Integer;
begin
  Len := Length(S);
  SetString(Result, PWideChar(S), Len);
  if Len > 0 then CharUpperBuffW(Pointer(Result), Len);
end;

function _WideLowerCase(const S: WideString): WideString;
var
  Len: Integer;
begin
  Len := Length(S);
  SetString(Result, PWideChar(S), Len);
  if Len > 0 then CharLowerBuffW(Pointer(Result), Len);
end;
{$ENDIF}

{$ENDIF}

{$IFDEF QI_UNICODE}
function QICompareStr(const S1, S2: WideString): Integer;
begin
{$IFDEF VCL6}
  Result := WideCompareStr(S1, S2);
{$ELSE}
  Result := _WideCompareStr(S1, S2);
{$ENDIF}
end;
{$ELSE}
function QICompareStr(const S1, S2: string): Integer;
begin
  Result := AnsiCompareStr(S1, S2);
end;
{$ENDIF}

{$IFDEF QI_UNICODE}
function QICompareText(const S1, S2: WideString): Integer;
begin
{$IFDEF VCL6}
  Result := WideCompareText(S1, S2);
{$ELSE}
  Result := _WideCompareText(S1, S2);
{$ENDIF}
end;
{$ELSE}
function QICompareText(const S1, S2: string): Integer;
begin
  Result := AnsiCompareText(S1, S2);
end;
{$ENDIF}

{$IFDEF QI_UNICODE}
function QIUpperCase(const S: WideString): WideString;
begin
{$IFDEF VCL6}
  Result := WideUpperCase(S);
{$ELSE}
  Result := _WideUpperCase(S);
{$ENDIF}
end;
{$ELSE}
function QIUpperCase(const S: string): string;
begin
  Result := AnsiUpperCase(S);
end;
{$ENDIF}

{$IFDEF QI_UNICODE}
function QILowerCase(const S: WideString): WideString;
begin
{$IFDEF VCL6}
  Result := WideLowerCase(S);
{$ELSE}
  Result := _WideLowerCase(S);
{$ENDIF}
end;
{$ELSE}
function QILowerCase(const S: string): string;
begin
  Result := AnsiLowerCase(S);
end;
{$ENDIF}

{$IFDEF QI_UNICODE}
function QIPos(const Substr, S: WideString): Integer;
var
  i: Integer;
  wc: WideChar;
  ws: WideString;
begin
  Result := 0;
  if (Substr = '') or (S = '') then Exit;

  for i := 1 to Length(S) do
  begin
    wc := S[i];
    if wc = Substr[1] then
    begin
      ws := Copy(S, i, Length(Substr));
      if QICompareStr(Substr, ws) = 0 then
      begin
        Result := i;
        Exit;
      end;
    end;
  end;
end;
{$ELSE}
function QIPos(const Substr, S: string): Integer;
begin
  Result := SysUtils.AnsiPos(Substr, S);
end;
{$ENDIF}

{$IFDEF QI_UNICODE}
procedure QIInsert(const Substr: WideString; var S: WideString; const Index: Integer);
begin
  System.Insert(Substr, S, Index);
end;
{$ELSE}
procedure QIInsert(const Substr: string; var S: string; const Index: Integer);
begin
  System.Insert(Substr, S, Index);
end;
{$ENDIF}

{$IFDEF QI_UNICODE}
procedure QIDelete(var S: WideString; const Index, Count: Integer);
begin
  System.Delete(S, Index, Count);
end;
{$ELSE}
procedure QIDelete(var S: string; const Index, Count: Integer);
begin
  System.Delete(S, Index, Count);
end;
{$ENDIF}

{$IFDEF QI_UNICODE}
{$IFNDEF VCL9}
procedure WideFmtStr(var Result: WideString; const Format: WideString;
  const Args: array of const);
const
  BufSize = 2048;
var
  Len, BufLen: Integer;
  Buffer: array[0..BufSize-1] of WideChar;
begin
  if Length(Format) < (BufSize - (BufSize div 4)) then
  begin
    BufLen := BufSize;
    Len := WideFormatBuf(Buffer, BufSize - 1, Pointer(Format)^, Length(Format), Args);
    if Len < BufLen - 1 then
    begin
      SetString(Result, Buffer, Len);
      Exit;
    end;
  end
  else
  begin
    BufLen := Length(Format);
    Len := BufLen;
  end;

  while Len >= BufLen - 1 do
  begin
    Inc(BufLen, BufLen);
    Result := '';          // prevent copying of existing data, for speed
    SetLength(Result, BufLen);
    Len := WideFormatBuf(Pointer(Result)^, BufLen - 1, Pointer(Format)^,
      Length(Format), Args);
  end;
  SetLength(Result, Len);
end;
{$ENDIF}
function QIFormat(const Format: WideString; const Args: array of const): WideString;
begin
  WideFmtStr(Result, Format, Args);
end;
{$ELSE}
function QIFormat(const Format: string; const Args: array of const): string;
begin
  Result := SysUtils.Format(Format, Args);
end;
{$ENDIF}

{$IFDEF QI_UNICODE}
{$IFNDEF VCL9}
function WideStringReplace(const S, OldPattern, NewPattern: Widestring;
  Flags: TReplaceFlags): Widestring;
var
  SearchStr, Patt, NewStr: Widestring;
  Offset: Integer;
begin
  if rfIgnoreCase in Flags then
  begin
    SearchStr := QIUpperCase(S);
    Patt := QIUpperCase(OldPattern);
  end else
  begin
    SearchStr := S;
    Patt := OldPattern;
  end;
  NewStr := S;
  Result := '';
  while SearchStr <> '' do
  begin
    Offset := Pos(Patt, SearchStr);
    if Offset = 0 then
    begin
      Result := Result + NewStr;
      Break;
    end;
    Result := Result + Copy(NewStr, 1, Offset - 1) + NewPattern;
    NewStr := Copy(NewStr, Offset + Length(OldPattern), MaxInt);
    if not (rfReplaceAll in Flags) then
    begin
      Result := Result + NewStr;
      Break;
    end;
    SearchStr := Copy(SearchStr, Offset + Length(Patt), MaxInt);
  end;
end;
{$ENDIF}
function QIStringReplace(const S, OldPattern, NewPattern: WideString;
  Flags: TReplaceFlags): WideString;
begin
  Result := WideStringReplace(S, OldPattern, NewPattern, Flags);
end;
{$ELSE}
function QIStringReplace(const S, OldPattern, NewPattern: string;
  Flags: TReplaceFlags): string;
begin
  Result := SysUtils.StringReplace(S, OldPattern, NewPattern, Flags);
end;
{$ENDIF}

{$IFDEF QI_UNICODE}
{$IFNDEF VCL10}
function WideQuotedStr(const S: WideString; Quote: WideChar): WideString;
var
  P, Src,
  Dest: PWideChar;
  AddCount: Integer;
begin
  AddCount := 0;
  P := WStrScan(PWideChar(S), Quote);
  while (P <> nil) do
  begin
    Inc(P);
    Inc(AddCount);
    P := WStrScan(P, Quote);
  end;

  if AddCount = 0 then
    Result := Quote + S + Quote
  else
  begin
    SetLength(Result, Length(S) + AddCount + 2);
    Dest := PWideChar(Result);
    Dest^ := Quote;
    Inc(Dest);
    Src := PWideChar(S);
    P := WStrScan(Src, Quote);
    repeat
      Inc(P);
      Move(Src^, Dest^, 2 * (P - Src));
      Inc(Dest, P - Src);
      Dest^ := Quote;
      Inc(Dest);
      Src := P;
      P := WStrScan(Src, Quote);
    until P = nil;
    P := WStrEnd(Src);
    Move(Src^, Dest^, 2 * (P - Src));
    Inc(Dest, P - Src);
    Dest^ := Quote;
  end;
end;
{$ENDIF}
function QIQuotedStr(const S: WideString; Quote: WideChar): WideString;
begin
  Result := WideQuotedStr(S, Quote);
end;
{$ELSE}
function QIQuotedStr(const S: string; Quote: Char): string;
begin
  Result := SysUtils.AnsiQuotedStr(S, Quote);
end;
{$ENDIF}

{$IFDEF QI_UNICODE}
{$IFNDEF VCL9}
function WideFormat(const Format: WideString; const Args: array of const): WideString;
begin
  WideFmtStr(Result, Format, Args);
end;

function WideReplaceStr(const AText, AFromText, AToText: WideString): WideString;
begin
  Result := WideStringReplace(AText, AFromText, AToText, [rfReplaceAll]);
end;

function WideReplaceText(const AText, AFromText, AToText: WideString): WideString;
begin
  Result := WideStringReplace(AText, AFromText, AToText, [rfReplaceAll, rfIgnoreCase]);
end;

function WStrAlloc(Size: Cardinal): PWideChar;
begin
  Size := SizeOf(Cardinal) + (Size * SizeOf(WideChar));
  GetMem(Result, Size);
  PCardinal(Result)^ := Size;
  Inc(PAnsiChar(Result), SizeOf(Cardinal));
end;

function WStrBufSize(const Str: PWideChar): Cardinal;
var
  P: PWideChar;
begin
  P := Str;
  Dec(PAnsiChar(P), SizeOf(Cardinal));
  Result := PCardinal(P)^ - SizeOf(Cardinal);
  Result := Result div SizeOf(WideChar);
end;
{$ENDIF}

{$IFNDEF VCL10}
function WStrMove(Dest: PWideChar; const Source: PWideChar; Count: Cardinal): PWideChar;
var
  Length: Integer;
begin
  Result := Dest;
  Length := Count * SizeOf(WideChar);
  Move(Source^, Dest^, Length);
end;
{$ENDIF}

{$IFNDEF VCL9}
function WStrNew(const Str: PWideChar): PWideChar;
var
  Size: Cardinal;
begin
  if Str = nil then Result := nil else
  begin
    Size := WStrLen(Str) + 1;
    Result := WStrMove(WStrAlloc(Size), Str, Size);
  end;
end;

procedure WStrDispose(Str: PWideChar);
begin
  if Str <> nil then
  begin
    Dec(PAnsiChar(Str), SizeOf(Cardinal));
    FreeMem(Str, Cardinal(Pointer(Str)^));
  end;
end;
{$ENDIF}

{$IFNDEF VCL9}
function WStrLen(Str: PWideChar): Cardinal;
begin
  Result := WStrEnd(Str) - Str;
end;

function WStrEnd(Str: PWideChar): PWideChar;
begin
  Result := Str;
  While Result^ <> #0 do
    Inc(Result);
end;
{$ENDIF}

{$IFNDEF VCL10}
function WStrCat(Dest: PWideChar; const Source: PWideChar): PWideChar;
begin
  Result := Dest;
  WStrCopy(WStrEnd(Dest), Source);
end;
{$ENDIF}

{$IFNDEF VCL9}
function WStrCopy(Dest, Source: PWideChar): PWideChar;
begin
  Result := WStrLCopy(Dest, Source, MaxInt);
end;

function WStrLCopy(Dest, Source: PWideChar; MaxLen: Cardinal): PWideChar;
var
  Count: Cardinal;
begin
  Result := Dest;
  Count := 0;
  while (Count < MaxLen) and (Source^ <> #0) do
  begin
    Dest^ := Source^;
    Inc(Source);
    Inc(Dest);
    Inc(Count);
  end;
  Dest^ := #0;
end;

function WStrPCopy(Dest: PWideChar; const Source: AnsiString): PWideChar;
begin
  Result := WStrPLCopy(Dest, Source, MaxInt);
end;

function WStrPLCopy(Dest: PWideChar; const Source: AnsiString; MaxLen: Cardinal): PWideChar;
begin
  Result := WStrLCopy(Dest, PWideChar(WideString(Source)), MaxLen);
end;
{$ENDIF}

{$IFNDEF VCL10}
function WStrScan(const Str: PWideChar; Chr: WideChar): PWideChar;
begin
  Result := Str;
  while Result^ <> Chr do
  begin
    if Result^ = #0 then
    begin
      Result := nil;
      Exit;
    end;
    Inc(Result);
  end;
end;

function WStrComp(Str1, Str2: PWideChar): Integer;
begin
  Result := WStrLComp(Str1, Str2, MaxInt);
end;

function WStrPos(Str, SubStr: PWideChar): PWideChar;
var
  PSave: PWideChar;
  P: PWideChar;
  PSub: PWideChar;
begin
  Result := nil;
  if (Str <> nil) and (Str^ <> #0) and (SubStr <> nil) and (SubStr^ <> #0) then
  begin
    P := Str;
    while P^ <> #0 do
    begin
      if P^ = SubStr^ then
      begin
        PSave := P;
        PSub := SubStr;
        while (P^ = PSub^) do
        begin
          Inc(P);
          Inc(PSub);
          if (PSub^ = #0) then
          begin
            Result := PSave;
            exit;
          end;
          if (P^ = #0) then
            Exit;
        end;
        P := PSave;
      end;
      Inc(P);
    end;
  end;
end;
{$ENDIF}

function WStrECopy(Dest, Source: PWideChar): PWideChar;
begin
  Result := WStrEnd(WStrCopy(Dest, Source));
end;

function WStrComp_EX(Str1, Str2: PWideChar; MaxLen: Cardinal; dwCmpFlags: Cardinal): Integer;
var
  Len1, Len2: Integer;
begin
  if MaxLen = Cardinal(MaxInt) then
  begin
    Len1 := -1;
    Len2 := -1;
  end
  else begin
    Len1 := Min(WStrLen(Str1), MaxLen);
    Len2 := Min(WStrLen(Str2), MaxLen);
  end;
  Result := CompareStringW(GetThreadLocale, dwCmpFlags, Str1, Len1, Str2, Len2) - 2;
end;

function WStrLComp(Str1, Str2: PWideChar; MaxLen: Cardinal): Integer;
begin
  Result := WStrComp_EX(Str1, Str2, MaxLen, 0);
end;

function WStrLIComp(Str1, Str2: PWideChar; MaxLen: Cardinal): Integer;
begin
  Result := WStrComp_EX(Str1, Str2, MaxLen, NORM_IGNORECASE);
end;

function WStrIComp(Str1, Str2: PWideChar): Integer;
begin
  Result := WStrLIComp(Str1, Str2, MaxInt);
end;

function WStrLower(Str: PWideChar): PWideChar;
begin
  Result := Str;
  CharLowerBuffW(Str, WStrLen(Str))
end;

function WStrUpper(Str: PWideChar): PWideChar;
begin
  Result := Str;
  CharUpperBuffW(Str, WStrLen(Str))
end;

function WStrRScan(const Str: PWideChar; Chr: WideChar): PWideChar;
var
  MostRecentFound: PWideChar;
begin
  if Chr = #0 then
    Result := WStrEnd(Str)
  else
  begin
    Result := nil;
    MostRecentFound := Str;
    while True do
    begin
      while MostRecentFound^ <> Chr do
      begin
        if MostRecentFound^ = #0 then
          Exit;
        Inc(MostRecentFound);
      end;
      Result := MostRecentFound;
      Inc(MostRecentFound);
    end;
  end;
end;

function WStrLCat(Dest: PWideChar; const Source: PWideChar; MaxLen: Cardinal): PWideChar;
begin
  Result := Dest;
  WStrLCopy(WStrEnd(Dest), Source, MaxLen - WStrLen(Dest));
end;

function WStrPas(const Str: PWideChar): WideString;
begin
  Result := Str;
end;

{$IFNDEF VCL10}
function WideLastChar(const S: WideString): PWideChar;
begin
  if S = '' then
    Result := nil
  else
    Result := @S[Length(S)];
end;
{$ENDIF}

{$IFNDEF VCL9}
function WideExtractQuotedStr(var Src: PWideChar; Quote: WideChar): Widestring;
var
  P, Dest: PWideChar;
  DropCount: Integer;
begin
  Result := '';
  if (Src = nil) or (Src^ <> Quote) then Exit;
  Inc(Src);
  DropCount := 1;
  P := Src;
  Src := WStrScan(Src, Quote);
  while Src <> nil do
  begin
    Inc(Src);
    if Src^ <> Quote then Break;
    Inc(Src);
    Inc(DropCount);
    Src := WStrScan(Src, Quote);
  end;
  if Src = nil then Src := WStrEnd(P);
  if ((Src - P) <= 1) then Exit;
  if DropCount = 1 then
    SetString(Result, P, Src - P - 1)
  else
  begin
    SetLength(Result, Src - P - DropCount);
    Dest := PWideChar(Result);
    Src := WStrScan(P, Quote);
    while Src <> nil do
    begin
      Inc(Src);
      if Src^ <> Quote then Break;
      Move(P^, Dest^, (Src - P) * SizeOf(WideChar));
      Inc(Dest, Src - P);
      Inc(Src);
      P := Src;
      Src := WStrScan(Src, Quote);
    end;
    if Src = nil then Src := WStrEnd(P);
    Move(P^, Dest^, (Src - P - 1) * SizeOf(WideChar));
  end;
end;
{$ENDIF}

{$IFNDEF VCL10}
function WideDequotedStr(const S: WideString; AQuote: WideChar): WideString;
var
  LText : PWideChar;
begin
  LText := PWideChar(S);
  Result := WideExtractQuotedStr(LText, AQuote);
  if Result = '' then
    Result := S;
end;
{$ENDIF}

{$ENDIF}


end.
