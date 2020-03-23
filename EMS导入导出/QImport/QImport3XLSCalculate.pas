unit QImport3XLSCalculate;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    System.SysUtils,
    System.Variants,
  {$ELSE}
    SysUtils,
    {$IFDEF VCL6}
      Variants,
    {$ENDIF}
  {$ENDIF}
  QImport3XLSFile;

function CalculateFormula(Cell: TbiffCell; ParsedData: PByteArray;
  DataSize: integer): Variant;

implementation

uses
  {$IFDEF VCL16}
    System.Math,
    System.Classes,
  {$ELSE}
    Math,
    Classes,
  {$ENDIF}
  QImport3XLSCommon,
  QImport3XLSConsts,
  QImport3XLSUtils;

type
  TxlsStack = class(TList)
  private
    function GetItems(Index: integer): Variant;
    procedure SetItems(Index: integer; Value: Variant);
  public
    function Add(Item: Variant): integer;
    procedure Clear; {$IFDEF VCL4}override;{$ENDIF}
    procedure Delete(Index: integer);
    destructor Destroy; override;
    procedure SetLength(Length: integer);

    property Items[Index: integer]: Variant read GetItems write SetItems; default;
  end;

  TxlsCalculator = class
  private
    FStack: TxlsStack;
    FPtr: integer;
    FCell: TbiffCell;
  public
    constructor Create(Cell: TbiffCell);
    destructor Destroy; override;

    procedure Push(Value: Variant);
    function Pop: Variant;
    function  Peek: Variant;
    procedure DoOperator(Operator: byte);
    procedure DoOptimizedFunction(ID: TXLS_FUNCTION_ID);
    function DoFunction(ID: TXLS_FUNCTION_ID): boolean;
    function DoVarFunction(ID: TXLS_FUNCTION_ID; ArgCount: integer): boolean;
  end;

{ TxlsStack }

function TxlsStack.Add(Item: Variant): integer;
var
  PV: PVariant; 
begin
  New(PV);
  PV^ := Item;
  Result := inherited Add(PV);
end;

procedure TxlsStack.Clear;
var
  i: integer;
begin
  for i := Count - 1 downto 0 do
    Delete(i);
  inherited;
end;

procedure TxlsStack.Delete(Index: integer);
begin
  Dispose(PVariant(inherited Items[Index]));
  inherited;
end;

destructor TxlsStack.Destroy;
begin
  Clear;
  inherited;
end;

function TxlsStack.GetItems(Index: integer): Variant;
begin
  Result := PVariant(inherited Items[Index])^;
end;

procedure TxlsStack.SetItems(Index: integer; Value: Variant);
begin
  PVariant(inherited Items[Index])^ := Value;
end;

procedure TxlsStack.SetLength(Length: integer);
var
  i: integer;
begin
  if Length <= Count then Exit;
  for i := Count to Length - 1 do
    Add(NULL);
end;

{ TxlsCalculator }

constructor TxlsCalculator.Create(Cell: TbiffCell);
begin
  inherited Create;
  FStack := TxlsStack.Create;
  FPtr := -1;
  FCell := Cell;
end;

destructor TxlsCalculator.Destroy;
begin
  FStack.Free;
  inherited;
end;

procedure TxlsCalculator.Push(Value: Variant);
begin
  Inc(FPtr);
  if FPtr >= FStack.Count then
    FStack.SetLength(FStack.Count + 32);
  FStack[FPtr] := Value;
end;

function TxlsCalculator.Pop: Variant;
begin
  if FPtr < 0 then begin
    Result := NULL;
    Exit;
    //raise ExlsFileError.Create(sEmptyStack);
  end;
  Result := FStack[FPtr];
  Dec(FPtr);
end;

function TxlsCalculator.Peek: Variant;
begin
  Result := FStack[FPtr];
end;

procedure TxlsCalculator.DoOperator(Operator: byte);
var
  Value: variant;
begin
  if FPtr < 1 then begin
    Value := NULL;
    Exit;
    //raise ExlsFileError.Create(sValueMissing);
  end;
  case Operator of
    ptgAdd   : Value := FStack[FPtr - 1] + FStack[FPtr];
    ptgSub   : Value := FStack[FPtr - 1] - FStack[FPtr];
    ptgMul   : Value := FStack[FPtr - 1] * FStack[FPtr];
    ptgDiv   : Value := FStack[FPtr - 1] / FStack[FPtr];
    ptgPower : Value := Power(FStack[FPtr - 1], FStack[FPtr]);
    ptgConcat:
      if (VarType(FStack[FPtr - 1]) = varString) and
         (VarType(FStack[FPtr]) = varString)
        then Value := FStack[FPtr - 1] + FStack[FPtr]
        else Value := BOOL_ERR_STR_VALUE;
    ptgLT    : Value := FStack[FPtr - 1] < FStack[FPtr];
    ptgLE    : Value := FStack[FPtr - 1] <= FStack[FPtr];
    ptgEQ    : Value := FStack[FPtr - 1] = FStack[FPtr];
    ptgGE    : Value := FStack[FPtr - 1] >= FStack[FPtr];
    ptgGT    : Value := FStack[FPtr - 1] > FStack[FPtr];
    ptgNE    : Value := FStack[FPtr - 1] <> FStack[FPtr];
    ptgUPlus : Value := FStack[FPtr] + 1;
    ptgUMinus: Value := -FStack[FPtr];
  end;
  Dec(FPtr);
  FStack[FPtr] := Value
end;

procedure TxlsCalculator.DoOptimizedFunction(ID: TXLS_FUNCTION_ID);
var
  RefArea, V: Variant;
  Col, Row, Sheet: integer;
  DataCell: TbiffCell;

  function DoOperation: Variant;
  begin
    if Assigned(DataCell) then
    begin
      if DataCell.IsFloat then begin
        if VarIsNull(V) then V := DataCell.AsFloat
        else
          case ID of
            fidSum: V := V + DataCell.AsFloat;
            fidMin: if DataCell.AsFloat < V then V := DataCell.AsFloat;
            fidMax: if DataCell.AsFloat > V then V := DataCell.AsFloat;
          end;
      end
    end else
      V := 0;
  end;

begin
  if not (ID in [fidSum, fidMin, fidMax]) then Exit;

  RefArea := Pop;
                                               
  V := NULL;
  if (VarType(RefArea) and varArray) = varArray then
    case VarArrayHighBound(RefArea, 1) of
      1: begin
        Col := RefArea[0];
        Row := RefArea[1];
        DataCell := FCell.WorkSheet.Cells[Row, Col];
        DoOperation;
      end;
      2: begin
        Col := RefArea[0];
        Row := RefArea[1];
        Sheet := RefArea[2];
        DataCell := FCell.WorkSheet.Workbook.WorkSheets[Sheet].Cells[Row, Col];
        DoOperation;
      end;
      3: begin
        for Col := RefArea[0] to RefArea[2] do
          for Row := RefArea[1] to RefArea[3] do begin
            DataCell := FCell.WorkSheet.Cells[Row, Col];
            DoOperation;
          end;
      end;
      4: begin
        Sheet := RefArea[4];
        for Col := RefArea[0] to RefArea[2] do
          for Row := RefArea[1] to RefArea[3] do begin
            DataCell := FCell.WorkSheet.Workbook.WorkSheets[Sheet].Cells[Row, Col];
            DoOperation;
          end;
      end;
    else 
      ExlsFileError.Create(sInvalidArraySize);
    end
  else
    if (VarType(RefArea) and varDouble) = varDouble then
      V := RefArea
    else
      raise ExlsFileError.Create(sInvalidAreaArgument);
  Push(V);
end;

function TxlsCalculator.DoFunction(ID: TXLS_FUNCTION_ID): boolean;
begin
  Result := true;
  case ID of
    fidCos: FStack[FPtr] := Cos(FStack[FPtr]);
    fidSin: FStack[FPtr] := Sin(FStack[FPtr]);
    fidTan: FStack[FPtr] := Tan(FStack[FPtr]);
    else Result := false;
  end;
end;

function TxlsCalculator.DoVarFunction(ID: TXLS_FUNCTION_ID;
  ArgCount: integer): boolean;
var
  R: Variant;

  procedure CallOptimizedFunction;
  var
    i: integer;
  begin
    R := 0;
    for i := 0 to ArgCount - 1 do begin
      if (VarType(Peek) and varArray) = varArray then
        DoOptimizedFunction(ID);

      case ID of
        fidSum: R := R + Pop;
        fidMax,
        fidMin: R := Pop;
      end;
    end;
  end;

  procedure DoCOUNT;
  var
    i: integer;
    V: Variant;
  begin
    R := 0;
    for i := 0 to ArgCount - 1 do begin
      if (VarType(Peek) and varArray) = varArray then begin
        V := Peek;
        case VarArrayHighBound(V, 1) of
          1, 2: R := R + 1;
          3, 4: R := R + (V[2] - V[0] + 1) * (V[3] - V[1] + 1);
          else raise ExlsFileError.Create(sInvalidArraySize);
        end;
      end
      else begin
        if VarType(Peek) in [varSmallint, varInteger, varSingle, varDouble,
          varCurrency, varDate, varBoolean {$IFDEF VCL6}, varShortInt, varWord,
          varInt64, varLongWord {$ENDIF}, varByte] then
          R := R + 1;
      end;
      Pop;
    end;
  end;

  procedure DoIF;
  var
    ResTrue, ResFalse: Variant;
  begin
    if ArgCount = 3
      then ResFalse := Pop
      else ResFalse := false;

    ResTrue := Pop;
    if Pop = true
      then R := ResTrue
      else R := ResFalse;
  end;

  procedure DoIsNA;
  begin
    Push((VarType(Peek) = varString) and (Pop = BOOL_ERR_STR_NA));
  end;

  procedure DoIsERROR;
  var
    i: integer;
  begin
    for i := 1 to High(BOOL_ERR_STRINGS) do begin
      if (VarType(Peek) = varString) and (Peek = BOOL_ERR_STRINGS[i]) then begin
        Pop;
        Push(true);
        Exit;
      end;
    end;
    Push(false);
  end;

  procedure DoAVERAGE;
  var
    i, Cnt: integer;
    V: Variant;
  begin
    R := 0;
    Cnt := 0;
    for i := 0 to ArgCount - 1 do begin
      if (VarType(Peek) and varArray) = varArray then begin
        V := Peek;
        case VarArrayHighBound(V, 1) of
          1, 2: Inc(Cnt);
          3, 4: Cnt := Cnt + (V[2] - V[0] + 1) * (V[3] - V[1] + 1);
          else raise Exception.Create(sInvalidArraySize);
        end;
        DoOptimizedFunction(fidSum);
        R := R + Peek;
      end
      else begin
        if VarType(Peek) in [varSmallint, varInteger, varSingle, varDouble,
          varCurrency, varDate, varBoolean, {$IFDEF VCL6}varShortInt, varWord,
          varLongWord, varInt64,{$ENDIF} varByte] then begin
          Inc(Cnt);
          R := R + Peek;
        end;
      end;
      Pop;
    end;
    R := R / Cnt;
  end;

begin
  Result := true;
  case ID of
    fidCount  : DoCOUNT;
    fidIf     : DoIF;
    fidIsNa   : DoIsNA;
    fidIsError: DoIsERROR;
    fidSum    : CallOptimizedFunction;
    fidAverage: DoAVERAGE;
    fidMax    : CallOptimizedFunction;
    fidMin    : CallOptimizedFunction;
    else Result := false;
  end;
  if Result then Push(R);
end;

function CalculateFormula(Cell: TbiffCell; ParsedData: PByteArray;
  DataSize: integer): Variant;
var
  Ptr: Pointer;
  Str: WideString;
  B: byte;
  Calculator: TxlsCalculator;
  C, R, i, j: integer;
  RefArea: Variant;
  DataCell: TbiffCell;

  procedure DecodeArea(ColIn, RowIn: integer; var ColOut, RowOut: integer);
  begin
    if (ColIn and $4000) = 0
      then ColOut := Shortint(ColIn and $FF)
      else ColOut := Cell.Col + Shortint(ColIn and $FF);
    if (ColIn and $8000) = 0
      then RowOut := RowIn
      else RowOut := Cell.Row + Smallint(RowIn);
  end;

  procedure CalculateFunction(ID, ArgCount: integer);
  var
    i: integer;
    ArrArgs: Variant;
    V: Variant;
    XLS: TxlsFile;
  begin
    ArrArgs := VarArrayCreate([0, ArgCount - 1], varVariant);
    for i := 0 to ArgCount - 1 do
      ArrArgs[ArgCount - 1 - i] := Calculator.Pop;

    V := NULL;

    XLS := Cell.WorkSheet.Workbook.ExcelFile;
    if Assigned(XLS.OnFunction) then
      XLS.OnFunction (XLS, XLS_FUNCTIONS[ID].Name, ArrArgs, V);

    if VarType(V) = NULL
      then Calculator.Push('[Unknown func ' + Str + ']')
      else Calculator.Push(V);
      
    VarArrayRedim(ArrArgs, 0);
  end;

begin
  Result := NULL;

  Calculator := TxlsCalculator.Create(Cell);
  try
    Ptr := ParsedData;
    while Integer(Ptr) - Integer(ParsedData) < DataSize do begin
      case Byte(Ptr^) of
        0: Break;
        ptgExp: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          Calculator.Push('<Shared>');
          //Break;
        end;
        ptgAdd..ptgPercent: begin
          Calculator.DoOperator(Byte(Ptr^));
          Ptr := Pointer(Integer(Ptr) + 1);
        end;
        ptgParen: Ptr := Pointer(Integer(Ptr) + 1);
        ptgMissArg: Ptr := Pointer(Integer(Ptr) + 1);
        ptgStr: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          B := Byte(Ptr^);
          Ptr := Pointer(Integer(Ptr) + 1);
          Str := ByteArrayToStr(Ptr, B);
          Ptr := Pointer(Integer(Ptr) + B);
        end;
        ptgAttr: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          if (Byte(Ptr^) and $04) = $04 then begin
            Ptr := Pointer(Integer(Ptr) + 1);
            Ptr := Pointer(Integer(Ptr) + (Word(Ptr^) + 2) * SizeOf(Word) - 3);
          end
          else if (Byte(Ptr^) and $10) = $10 then
            Calculator.DoOptimizedFunction(fidSum);
          Ptr := Pointer(Integer(Ptr) + 3);
        end;
        ptgSheet: Ptr := Pointer(Integer(Ptr) + 11);
        ptgEndSheet: Ptr := Pointer(Integer(Ptr) + 5);
        ptgErr: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          Calculator.Push(ErrCodeToString(Byte(Ptr^)));
          Ptr := Pointer(Integer(Ptr) + 1);
        end;
        ptgBool: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          if Byte(Ptr^) = 0
            then Calculator.Push(false)
            else Calculator.Push(true);
          Ptr := Pointer(Integer(Ptr) + 1);
        end;
        ptgInt: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          Calculator.Push(Smallint(Ptr^));
          Ptr := Pointer(Integer(Ptr) + 2);
        end;
        ptgNum: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          Calculator.Push(Double(Ptr^));
          Ptr := Pointer(Integer(Ptr) + 8);
        end;
        ptgRef, ptgRefV, ptgRefA: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          with PBIFF_PTGRef(Ptr)^ do begin
            DataCell := Cell.WorkSheet.Cells[Row, Col and $3FFF];
            if Assigned(DataCell) and DataCell.IsFloat
              then Calculator.Push(DataCell.AsFloat)
              else Calculator.Push(0);
            Ptr := Pointer(Integer(Ptr) + SizeOf(TBIFF_PTGRef));
          end;
        end;
        ptgRefN, ptgRefNV, ptgRefNA: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          with PBIFF_PTGRef(Ptr)^ do begin
            DecodeArea(Cell.Col, Cell.Row, C, R);
            Calculator.Push(VarArrayOf([C, R]));
            Ptr := Pointer(Integer(Ptr) + SizeOf(TBIFF_PTGRef));
          end;
        end;
        ptgArea, ptgAreaV, ptgAreaA: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          with PBIFF_PTGArea(Ptr)^ do begin
            Calculator.Push(VarArrayOf([Col1 and $3FFF, Row1, Col2 and $3FFF, Row2]));
            Ptr := Pointer(Integer(Ptr) + SizeOf(TBIFF_PTGArea));
          end;
        end;
        ptgAreaN, ptgAreaNV, ptgAreaNA: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          with PBIFF_PTGArea(Ptr)^ do begin
            DecodeArea(Col1, Row1, C, R);
            RefArea := VarArrayOf([C, R, 0, 0]);
            DecodeArea(Col2, Row2, C, R);
            RefArea[2] := C;
            RefArea[3] := R;
          end;
          Calculator.Push(RefArea);
        end;
        ptgRefErr, ptgRefErrV, ptgRefErrA: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          Ptr := Pointer(Integer(Ptr) + 1);
          Calculator.Push(BOOL_ERR_STR_REF);
        end;
        ptgAreaErr, ptgAreaErrV, ptgAreaErrA: begin
          Ptr := Pointer(Integer(Ptr) + 7);
          Ptr := Pointer(Integer(Ptr) + 2);
          Calculator.Push(BOOL_ERR_STR_REF);
        end;
        ptgName, ptgNameV, ptgNameA: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          with PBIFF_PTGName(Ptr)^ do begin
            Calculator.Push(Cell.WorkSheet.Workbook.Globals.NameList[NameIndex - 1].Name);
            Ptr := Pointer(Integer(Ptr) + SizeOf(TBIFF_PTGName));
          end;
        end;
        {ptgNameX, ptgNameXV, ptgNameXA: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          with PBIFF_PTGNameX(Ptr)^ do begin
            Calculator.Push(XLS.GetName(ntExternName,ExtSheet,NameIndex));
            P := Pointer(Integer(P) + SizeOf(TPTGNameX8));
          end;
        end;}
        ptgArray, ptgArrayV, ptgArrayA: begin
          Ptr := Pointer(Integer(Ptr) + 1);
  //        P2 := Pointer(Integer(PBuf) + PRecFORMULA(PBuf)^.ParseLen);


  //        S := GetArray;
  //        Stack.Add(Copy(S,1,Length(S) - 1) + '}');
  //        P := Pointer(Integer(P) + 7);
        end;
        ptgFunc, ptgFuncV, ptgFuncA: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          if Word(Ptr^) <= High(XLS_FUNCTIONS) then begin
            if not Calculator.DoFunction(TXLS_FUNCTION_ID(Word(Ptr^))) then
              CalculateFunction(Byte(Ptr^), XLS_FUNCTIONS[Byte(Ptr^)].Min);
            Ptr := Pointer(Integer(Ptr) + 1);
          end
          else begin
            Calculator.Push('?<' + IntToStr(Word(Ptr^)) + '>?');
            Ptr := Pointer(Integer(Ptr) + 1);
          end;
          Ptr := Pointer(Integer(Ptr) + 1);
        end;
        ptgFuncVar, ptgFuncVarV, ptgFuncVarA: begin
          Ptr := Pointer(Integer(Ptr) + 1);
          i := Byte(Ptr^) and $7F;
          j := -1;
          Ptr := Pointer(Integer(Ptr) + 1);
          if Word(Ptr^) <= High(XLS_FUNCTIONS) then
            j := Word(Ptr^) and $7FFF;

          if j <= Integer(High(TXLS_FUNCTION_ID))
            then Calculator.DoVarFunction(TXLS_FUNCTION_ID(j), i)
            else CalculateFunction(j,i);

          Ptr := Pointer(Integer(Ptr) + 1);
          Ptr := Pointer(Integer(Ptr) + 1);
        end;
        ptgRef3d, ptgRef3dV, ptgRef3dA, ptgRefErr3d,
        ptgRefErr3dV, ptgRefErr3dA: begin
          //i := Byte(Ptr^);
          Ptr := Pointer(Integer(Ptr) + 1);
          with PBIFF_PTGRef3D(Ptr)^ do begin
            {if i in [ptgRefErr3d, ptgRefErr3dV]
              then Str := '!#REF!'
              else Calculator.Push(XLS.GetNameValue(ntExternSheet,Index,Col and $3FFF,Row));}

            Ptr := Pointer(Integer(Ptr) + SizeOf(TBIFF_PTGRef3D));
          end
        end;
        {ptgArea3d, ptgArea3dV, ptgArea3dA, ptgAreaErr3d,
        ptgAreaErr3dV, ptgAreaErr3dA: begin
          i := Byte(Ptr^);
          Ptr := Pointer(Integer(Ptr) + SizeOf(TBIFF_PTGRef3D));
          with PBIFF_PTGArea3D(Ptr)^ do begin
            if Index < Cell.WorkSheet.Workbook.Sheets.Count
              then Str := Cell.WorkSheet.Workbook.Sheets[Index]. XLS.Sheets[Index].Name
            else
              S := '<??? Ref3d>';
            if i in [ptgAreaErr3d,ptgAreaErr3dV] then
              S := S + '!#REF!'
            else
              VStack.Push(VarArrayOf([Col1 and $3FFF,Row1,Col2 and $3FFF,Row2,Index]));
            P := Pointer(Integer(P) + SizeOf(TPTGArea3d8));
          end
          else with PPTGArea3d7(P)^ do begin
            if Smallint(SheetIndex) >= 0 then
              S := '[EXTERN ???]:'
            else
              S := '';
            if (IndexFirst = $FFFF) or (IndexLast = $FFFF) then
              S := S + '[DELETED]'
            else if IndexFirst = IndexLast then
              S := S + XLS.Sheets[IndexLast].Name
            else
              S := S + XLS.Sheets[IndexFirst].Name + ':' + XLS.Sheets[IndexLast].Name;
            if i in [ptgAreaErr3d,ptgAreaErr3dV] then
              VStack.Push(S + '!#REF!')
            else
              VStack.Push(VarArrayOf([Col1,Row1 and $3FFF,Col2,Row2 and $3FFF,SheetIndex]));
            P := Pointer(Integer(P) + SizeOf(TPTGArea3d7));
          end;
        end;}
        else begin
          Calculator.Push(Format(sUnknownPTG, [Byte(Ptr^)]));
          Break;
        end;
      end;
    end;
    Result := Calculator.Pop;
  finally
    Calculator.Free;
  end;
end;

end.
