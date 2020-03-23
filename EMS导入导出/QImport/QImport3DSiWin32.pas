unit QImport3DSiWin32;
{$I QImport3VerCtrl.Inc}
interface

uses
  {$IFDEF VCL16}
    System.SysUtils,
    Winapi.Windows;
  {$ELSE}
    SysUtils,
    Windows;
  {$ENDIF}

function DSiFileExistsW(const fileName: WideString): Boolean;
function DSiIsFileCompressedW(const fileName: WideString): Boolean;
function DSiIsFileCompressed(const fileName: string): Boolean;
function DSiCompressFile(fileHandle: THandle): Boolean;

implementation

const
  COMPRESSION_FORMAT_DEFAULT = 1;
  FILE_DEVICE_FILE_SYSTEM = 9;
  METHOD_BUFFERED = 0;
  FILE_READ_ACCESS = 1;
  FILE_WRITE_ACCESS = 2;
  FSCTL_SET_COMPRESSION = (FILE_DEVICE_FILE_SYSTEM shl 16) or
                          ((FILE_READ_ACCESS or FILE_WRITE_ACCESS) shl 14) or
                          (16 shl 2) or (METHOD_BUFFERED);

function DSiSetCompression(fileHandle: THandle; compressionFormat: integer): boolean;
var
  comp: SHORT;
  res: DWORD;
begin
  if Win32Platform <> VER_PLATFORM_WIN32_NT then //only NT can compress files
    Result := true
  else begin
    res := 0;
    comp := compressionFormat;
    Result := DeviceIoControl(fileHandle, FSCTL_SET_COMPRESSION, @comp, SizeOf(SHORT),
      nil, 0, res, nil);
  end;
end;

function DSiFileExistsW(const fileName: WideString): Boolean;
var
  code: Integer;
begin
  code := GetFileAttributesW(PWideChar(fileName));
  Result := (code <> -1) and ((FILE_ATTRIBUTE_DIRECTORY and code) = 0);
end;

function DSiIsFileCompressedW(const fileName: WideString): Boolean;
begin
  if Win32Platform <> VER_PLATFORM_WIN32_NT then //only NT can compress files
    Result := false
  else
    Result := (GetFileAttributesW(PWideChar(fileName)) and FILE_ATTRIBUTE_COMPRESSED) =
      FILE_ATTRIBUTE_COMPRESSED;
end;

function DSiIsFileCompressed(const fileName: string): Boolean;
begin
  if Win32Platform <> VER_PLATFORM_WIN32_NT then //only NT can compress files
    Result := false
  else
    Result := (GetFileAttributes(PChar(fileName)) and FILE_ATTRIBUTE_COMPRESSED) =
      FILE_ATTRIBUTE_COMPRESSED;
end;

function DSiCompressFile(fileHandle: THandle): Boolean;
begin
  Result := DSiSetCompression(fileHandle, COMPRESSION_FORMAT_DEFAULT);
end;

end.
