program ProcJMJS;

uses
  Forms,
  fmProcJMJS in 'fmProcJMJS.pas' {fmJMJS};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmJMJS, fmJMJS);
  Application.Run;
end.
