program fmDecrypt_PL;

uses
  Forms,
  Decrypt_PL in 'Decrypt_PL.pas' {fmProcDecrypt};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmProcDecrypt, fmProcDecrypt);
  Application.Run;
end.
