unit TestMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RzButton, StdCtrls, RzEdit, Mask, RzLabel, ExtCtrls, RzPanel;

type
  TForm1 = class(TForm)
    RzPanel1: TRzPanel;
    RzLabel1: TRzLabel;
    edtMM: TRzEdit;
    RzLabel2: TRzLabel;
    edtRL: TRzEdit;
    RzMemo1: TRzMemo;
    RzLabel3: TRzLabel;
    RzButton1: TRzButton;
    procedure RzButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure InitPChar(var pPch: PChar; iLen: Integer);
  end;

var
  Form1: TForm1;

function GetDecrypt_DES(pCODE:PChar;var pREAL,pMsg:PChar):boolean;stdcall;external 'Decrypt.dll' name 'GetDecrypt_DES';

implementation

{$R *.dfm}

procedure TForm1.InitPChar(var pPch: PChar; iLen: Integer);
begin
  pPch := stralloc(iLen);
end;

procedure TForm1.RzButton1Click(Sender: TObject);
var
  sMSG,sREAL:Pchar;
  bSecc:Boolean;
begin
  InitPChar(sMSG, 100);
  InitPChar(sREAL, 100);
  bSecc := GetDecrypt_DES(Pchar(edtMM.Text),sREAL,sMSG);
  if bSecc then
    edtRL.Text:= sREAL
  else
    edtRL.Text:='';
  RzMemo1.Clear;
  RzMemo1.Lines.Add(sMSG);
end;

end.
