unit fmProcIDVerification;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, RzLabel, ExtCtrls, RzPanel, RzButton, Mask, RzEdit;

type
  TForm1 = class(TForm)
    RzPanel1: TRzPanel;
    RzLabel1: TRzLabel;
    edtSFZ: TRzEdit;
    RzButton1: TRzButton;
    RzButton2: TRzButton;
    procedure RzButton1Click(Sender: TObject);
    procedure RzButton2Click(Sender: TObject);
  private
    { Private declarations }
    function IsValidSFZ(psSFZ: string): Boolean;
    function OldSFZToNew(pOLD: string): string;
    function GetDDD(sNR:string):string;
    function ReturnDDD(sNR:string):string;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

function TForm1.GetDDD(sNR: string): string;
begin
  if length(sNR)<>17 then exit;
  result:=ReturnDDD(sNR);
end;

function TForm1.IsValidSFZ(psSFZ: string): Boolean;
begin
  Result := psSFZ = OldSFZToNew(Copy(psSFZ, 1, 17));
end;

function TForm1.OldSFZToNew(pOLD: string): string;
const
  W: array[1..18] of integer = (7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2, 1);
  A: array[0..10] of char = ('1', '0', 'x', '9', '8', '7', '6', '5', '4', '3', '2');
var
  i, j, S: integer;
begin
  if Length(pOLD) <> 17 then
    Result := ''
  else
  begin
    Result := pOLD;
    //Insert(copy(edtSFZ.Text,7,2), Result, 7);
    S := 0;
    try
      for i := 1 to 17 do
      begin
        j := StrToInt(Result[i]) * W[i];
        S := S + j;
      end;
    except
      result := '';
      exit;
    end;
    S := S mod 11;
    Result := Result + A[S];
  end;
end;

function TForm1.ReturnDDD(sNR: string): string;
const
  W: array[1..18] of integer = (7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2, 1);
  A: array[0..10] of char = ('1', '0', 'x', '9', '8', '7', '6', '5', '4', '3', '2');
var
  i, j, S: integer;
begin
  if Length(sNR) <> 17 then
    Result := ''
  else
  begin
    Result := sNR;
    S := 0;
    try
      for i := 1 to 17 do
      begin
        j := StrToInt(Result[i]) * W[i];
        S := S + j;
      end;
    except
      result := '';
      exit;
    end;
    S := S mod 11;
    Result := A[S];
  end;
end;

procedure TForm1.RzButton1Click(Sender: TObject);
begin
  if edtSFZ.Text='' then
  begin
    showmessage('请输入身份证号');
    exit;
  end;
  if not IsValidSFZ(AnsiLowerCase(edtSFZ.Text)) then
  begin
    showmessage('身份证号码验证有误！');
    exit;
  end
  else
  begin
    showmessage('身份证号码验证成功！');
  end;
end;

procedure TForm1.RzButton2Click(Sender: TObject);
begin
  showmessage(GetDDD(edtSFZ.Text));
end;

end.
