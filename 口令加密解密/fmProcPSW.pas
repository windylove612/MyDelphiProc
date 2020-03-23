unit fmProcPSW;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RzButton, StdCtrls, Mask, RzEdit, RzLabel,unEncrypt, DB,
  DBTables;

type
  TForm1 = class(TForm)
    RzLabel1: TRzLabel;
    RzLabel2: TRzLabel;
    RzEdit1: TRzEdit;
    RzEdit2: TRzEdit;
    RzButton1: TRzButton;
    RzButton2: TRzButton;
    procedure RzButton1Click(Sender: TObject);
    procedure RzButton2Click(Sender: TObject);
  private
    { Private declarations }
    pPaswd		:Pchar;
    paswdDataSize:integer;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.RzButton1Click(Sender: TObject);
var
  tmpstr,mpstr	:string;
  i		:integer;
  s:TVarBytesField;
begin
  if RzEdit1.Text='' then
  begin
    showmessage('请输入明文！');
    exit;
  end;
  paswdDataSize:=12;
  GetMem(pPaswd,paswdDataSize+4);
  FillChar(pPaswd^,paswdDatasize+4,$0);

  tmpstr := RzEdit1.Text;
  for i := Length(RzEdit1.Text)+1 to 12 do
	  tmpstr := tmpstr + ' ';
  Move(pChar(Encrypt(tmpstr))^, (pPaswd+2)^, 12);
  PSmallInt(pPaswd)^ := 12;
  RzEdit2.Text := Decrypt(StrPas(pPaswd+2));
  //RzEdit2.Text := pPaswd;
end;

procedure TForm1.RzButton2Click(Sender: TObject);
var
  DataSize,MyII      :integer;
  BinPassword	:string;
  i,j			:integer;
  iCount          :integer;
  S               :string;
  pPassword       :PChar;
  sPassword       :string;
  iSize           :integer;
  sTMP            :PChar;
begin
  if RzEdit2.Text='' then
  begin
    showmessage('请输入暗文！');
    exit;
  end;
  paswdDataSize:=12;
  GetMem(pPaswd,paswdDataSize+4);
  FillChar(pPaswd^,paswdDatasize+4,$0);

  pPaswd:=TVarBytesField(RzEdit2.Text);
  RzEdit1.Text:=Decrypt(StrPas(pPaswd+2));
end;

end.
