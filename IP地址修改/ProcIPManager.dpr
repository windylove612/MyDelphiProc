program ProcIPManager;

uses
  Forms,
  fmProcIPManager in 'fmProcIPManager.pas' {fmIPManager},
  LGetAdapterInfo in 'LGetAdapterInfo.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmIPManager, fmIPManager);
  Application.Run;
end.
