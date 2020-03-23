program ProcRecord;

uses
  Forms,
  fmProcRecord in 'fmProcRecord.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
