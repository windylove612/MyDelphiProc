unit QImport3StrTypes;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    System.WideStrings,
    {$IFNDEF QI_UNICODE}
      System.Classes,
    {$ENDIF}
    {$IFDEF WIN64}
      Vcl.Grids,
    {$ENDIF}
  {$ELSE}
    {$IFDEF VCL10}
      WideStrings,
    {$ELSE}
      QImport3WideStrings,
    {$ENDIF}
    {$IFNDEF QI_UNICODE}
      Classes,
      Grids,
    {$ENDIF}
  {$ENDIF}
  QImport3WideStringGrid,
  QImport3GpTextFile;

type
 {$IFDEF QI_UNICODE}
  qiString = WideString;
  qiChar = WideChar;
  PqiChar = PWideChar;
  TqiStrings = TWideStrings;
  TqiStringList = TWideStringList;
  TqiStringGrid = {$IFDEF WIN64}TStringGrid{$ELSE}TEmsWideStringGrid{$ENDIF};
  TqiTextFile = TGpTextFile;
 {$ELSE}
  qiString = string;
  qiChar = Char;
  PqiChar = PChar;
  TqiStrings = TStrings;
  TqiStringList = TStringList;
  TqiStringGrid = TStringGrid;
  TqiTextFile = TextFile;
 {$ENDIF}

implementation

end.
