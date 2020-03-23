program fmCSRQ;

uses
  Forms,
  CSRQ in 'CSRQ.pas' {fmProcCSRQ_SFZ};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmProcCSRQ_SFZ, fmProcCSRQ_SFZ);
  Application.Run;
end.
