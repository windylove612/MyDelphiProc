program TestDLL;

uses
  Forms,
  TestMain in 'TestMain.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
