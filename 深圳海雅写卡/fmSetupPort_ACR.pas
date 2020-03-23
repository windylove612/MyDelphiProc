unit fmSetupPort_ACR;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RzButton, BMPBTN, ExtCtrls, RzPanel, RzRadGrp,Registry;

type
  TUSBPortSetting = record
    iPort	:Integer;
  end;
  
  TSetPort_ACR = class(TForm)
    grpBaud: TRzRadioGroup;
    btnOk: TBmpBtn;
    btnCancel: TBmpBtn;
    procedure btnOkClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
  private
    { Private declarations }
    pData	:^TUSBPortSetting;
  public
    { Public declarations }
    constructor Create(AOwner:TComponent; var Data:TUSBPortSetting);
  end;

var
  SetPort_ACR: TSetPort_ACR;
  function DoSetPort_ACR(AOwner:TComponent; var Data:TUSBPortSetting):Boolean;
  procedure DoGetDefPortSetting_ACR(var Data:TUSBPortSetting);

implementation

{$R *.dfm}
const
  KeyValue = 'SOFTWARE\BF\3.1\ACRPORTSETTING';
  FldPort = 'Port';

procedure DoGetDefPortSetting_ACR(var Data:TUSBPortSetting);
var
  iVal	:integer;
  bVal	:boolean;
begin
  Data.iPort := 1;
  with TRegistry.Create do
  begin
    RootKey := HKEY_LOCAL_MACHINE;
    //if OpenKeyReadOnly(KeyValue) then
    if OpenKey(KeyValue, False) then
    begin
	try
	  iVal := ReadInteger(FldPort);
	  Data.iPort := iVal;
	except
	end;
	CloseKey();
    end
  end;
end;

function DoSetPort_ACR(AOwner:TComponent; var Data:TUSBPortSetting):Boolean;
var
  SetPort_ACR: TSetPort_ACR;
begin
  SetPort_ACR := TSetPort_ACR.Create(AOwner, Data);
  Result := (SetPort_ACR.ShowModal=mrOk);
  SetPort_ACR.Free;
end;

{ TSetPort_ACR }

constructor TSetPort_ACR.Create(AOwner: TComponent;
  var Data: TUSBPortSetting);
begin
  inherited Create(AOwner);

  pData := @Data;
  with Data do
  begin
    grpBaud.ItemIndex := iPort;
  end;
end;

procedure TSetPort_ACR.btnOkClick(Sender: TObject);
begin
  with pData^ do
  begin
    iPort := grpBaud.ItemIndex;
  end;
  with TRegistry.Create do
  begin
    RootKey := HKEY_LOCAL_MACHINE;
    //if OpenKeyReadOnly(KeyValue) then
    if OpenKey(KeyValue, True) then
    begin
	try
	  WriteInteger(FldPort, pData^.iPort);
	except
	end;
	CloseKey();
    end
  end;
  ModalResult := mrOk;
end;

procedure TSetPort_ACR.btnCancelClick(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

end.
