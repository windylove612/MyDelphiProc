unit UnitMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TMainForm = class(TForm)
    GetMac:      TButton;
    Memo: TMemo;
    procedure GetMacClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

uses
  Registry;

{$R *.dfm}

function GetNetCardMac(NetCardName: string; Current: Boolean): string; inline;
const
  OID_802_3_PERMANENT_ADDRESS: Integer = $01010101;
  OID_802_3_CURRENT_ADDRESS: Integer = $01010102;
  IOCTL_NDIS_QUERY_GLOBAL_STATS: Integer = $00170002;
var
  hDevice: THandle;
  inBuf: Integer;
  outBuf: array[1..256] of Byte;
  BytesReturned: DWORD;
  MacAddr: string;
  i: integer;
begin
  if Current then
    inBuf  := OID_802_3_CURRENT_ADDRESS
  else
    inBuf  := OID_802_3_PERMANENT_ADDRESS;
  hDevice := 0;
  Result := '';
  try
    hDevice := CreateFile(PChar('\\.\' + NetCardName), GENERIC_READ or
      GENERIC_WRITE, FILE_SHARE_READ or FILE_SHARE_WRITE, nil, OPEN_EXISTING, 0, 0);
    if hDevice <> INVALID_HANDLE_VALUE then
    begin
      if DeviceIoControl(hDevice, IOCTL_NDIS_QUERY_GLOBAL_STATS, @inBuf, 4,
        @outBuf, 256, BytesReturned, nil) then
      begin
        MacAddr := '';
        for i := 1 to BytesReturned do
        begin
          MacAddr := MacAddr + IntToHex(outbuf[i], 2);
        end;
        Result := MacAddr;
      end;
    end;
  finally
    if not hDevice <> INVALID_HANDLE_VALUE then
      CloseHandle(hDevice);
  end;
end;

function GetNetCardVendor(NetCardName: string): string; inline;
const
  OID_GEN_VENDOR_DESCRIPTION : Integer = $0001010D;
  IOCTL_NDIS_QUERY_GLOBAL_STATS: Integer = $00170002;
var
  hDevice: THandle;
  inBuf: Integer;
  outBuf: array[1..256] of AnsiChar;
  BytesReturned: DWORD;
begin
  inBuf  := OID_GEN_VENDOR_DESCRIPTION;
  hDevice := 0;
  Result := '';
  try
    hDevice := CreateFile(PChar('\\.\' + NetCardName), GENERIC_READ or
      GENERIC_WRITE, FILE_SHARE_READ or FILE_SHARE_WRITE, nil, OPEN_EXISTING, 0, 0);
    if hDevice <> INVALID_HANDLE_VALUE then
    begin
      if DeviceIoControl(hDevice, IOCTL_NDIS_QUERY_GLOBAL_STATS, @inBuf, 4,
        @outBuf, 256, BytesReturned, nil) then
      Result := string(AnsiString(outBuf));
    end;
  finally
    if not hDevice <> INVALID_HANDLE_VALUE then
      CloseHandle(hDevice);
  end;
end;


procedure TMainForm.GetMacClick(Sender: TObject);
const
  REG_NWC = '\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkCards';
var
  NetWorkCards: TStringList;
  ServiceName: string;
  i: Integer;
begin
  Memo.Clear;
  with TRegistry.Create do
  begin
    RootKey := HKEY_LOCAL_MACHINE;
    NetWorkCards := TStringList.Create;
    try
      if OpenKeyReadOnly(REG_NWC) then
        GetKeyNames(NetWorkCards);
      for i := 0 to NetWorkCards.Count - 1 do
      begin
        if OpenKeyReadOnly(REG_NWC + '\' + NetWorkCards[i]) then
        begin
          ServiceName := ReadString('ServiceName');
          Memo.Lines.Add('描述: ' + ReadString('Description'));
          Memo.Lines.Add('永久地址: ' + GetNetCardMac(ServiceName, False));
          Memo.Lines.Add('当前地址: ' + GetNetCardMac(ServiceName, True));
          Memo.Lines.Add('厂商: ' + GetNetCardVendor(ServiceName));
          Memo.Lines.Add('');
          Memo.Lines.Add('');
          Memo.Lines.Add('');
        end;
      end;
    finally
      NetWorkCards.Free;
    end;
  end;
end;

end.

