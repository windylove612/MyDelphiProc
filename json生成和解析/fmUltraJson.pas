unit fmUltraJson;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RzButton, StdCtrls, RzEdit,superobject,superxmlparser, Mask;

type
  TForm1 = class(TForm)
    RzMemo1: TRzMemo;
    RzButton1: TRzButton;
    RzButton2: TRzButton;
    RzMemo2: TRzMemo;
    RzButton3: TRzButton;
    RzEdit1: TRzEdit;
    RzButton4: TRzButton;
    RzEdit2: TRzEdit;
    RzEdit3: TRzEdit;
    procedure RzButton2Click(Sender: TObject);
    procedure RzButton1Click(Sender: TObject);
    procedure RzButton3Click(Sender: TObject);
    procedure RzButton4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.RzButton2Click(Sender: TObject);
begin
  RzMemo1.Clear;
end;

procedure TForm1.RzButton1Click(Sender: TObject);
var
  obj: ISuperObject;
  str: string;
  arr: TSuperArray;
  item: TSuperAvlEntry;
  i:integer;
begin
  obj := SO(RzMemo1.Lines.Text);
  if RzEdit1.Text='' then
  begin
    RzMemo2.Lines.add(obj.AsString);
    ShowMessage(obj.AsString);
  end
  else
  begin
    RzMemo2.Lines.add(obj[RzEdit1.Text].AsString);
    ShowMessage(obj[RzEdit1.Text].AsString);
  end;
end;

procedure TForm1.RzButton3Click(Sender: TObject);
begin
  RzMemo2.Clear;
end;

procedure TForm1.RzButton4Click(Sender: TObject);
var
  obj: ISuperObject;
  str: string;
  arr: TSuperArray;
  item: ISuperObject;
  i:integer;
begin
  obj := SO(RzMemo1.Lines.Text);
  arr := obj[RzEdit2.text].AsArray;
  for i:=0 to arr.Length-1 do
  begin
    RzMemo2.Lines.add(arr[i].AsString);
    ShowMessage(arr[i].AsString);
  end;
end;

end.
