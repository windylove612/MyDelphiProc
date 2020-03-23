unit QImport3JScriptEngine;
{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    System.Classes,
  {$ELSE}
    Classes,
  {$ENDIF}
  QImport3ScriptEngine,
  QImport3StrTypes,
  QImport3MSScriptControlTLB;

type
  TQImport3JScriptEngine = class(TQImport3ScriptEngine)
  private
    FScriptControl: TScriptControl;
  protected
    function RunScript(const ScriptText: qiString): Variant; override;
    function GetError: TScriptErrorMsg; override;
    function QuoteStringValue(const Value: qiString): qiString; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

implementation

{ TQImport3JScriptEngine }

constructor TQImport3JScriptEngine.Create(AOwner: TComponent);
const
  DefLanguage = 'JScript';
begin
  inherited Create(AOwner);
  FScriptControl := TScriptControl.Create(nil);
  FScriptControl.Language := DefLanguage;
end;

destructor TQImport3JScriptEngine.Destroy;
begin
  FScriptControl.Free;
  inherited;
end;

function TQImport3JScriptEngine.GetError: TScriptErrorMsg;
begin
  Result.Line := FScriptControl.Error.Line;
  Result.Column := FScriptControl.Error.Column;
  Result.ErrorCode := FScriptControl.Error.Number;
  Result.ErrorDescription := FScriptControl.Error.Description;
  Result.SourceText := FScriptControl.Error.Text;
  FScriptControl.Error.Clear;
end;


function TQImport3JScriptEngine.QuoteStringValue(
  const Value: qiString): qiString;
const
  QuoteChar = '"';
begin
  Result := QuoteStringValue(Value, QuoteChar)
end;

function TQImport3JScriptEngine.RunScript(const ScriptText: qiString): Variant;
begin
  Result := FScriptControl.Eval(ScriptText);
end;

end.
