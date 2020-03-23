program ProcWriteICK;

uses
  Forms,
  fmProcWriteICK in 'fmProcWriteICK.pas' {ProcZK};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TProcZK, ProcZK);
  Application.Run;
end.
