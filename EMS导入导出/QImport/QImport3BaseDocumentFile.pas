unit QImport3BaseDocumentFile;

{$I QImport3VerCtrl.Inc}

interface

uses
  QImport3StrTypes;

type
  TBaseDocumentFile = class
  private
    FFileName: qiString;
    FLoaded: boolean;
    FCurrFolder: qiString;
    procedure SetFileName(const Value: qiString);
  protected
    procedure LoadXML(WorkDir: qiString); virtual; abstract;
  public
    constructor Create; virtual;

    procedure Decompress;
    procedure DeleteTempFolder;
    procedure Load;
    procedure Clear;

    property Loaded: boolean read FLoaded;
    property FileName: qiString read FFileName write SetFileName;
  end;

implementation

uses
  {$IFDEF VCL16}
    System.SysUtils,
    System.Classes,
    Winapi.Windows,
    Vcl.StdCtrls,
    Vcl.ExtCtrls,
    {$IFDEF VER130}
      Vcl.FileCtrl,
    {$ENDIF}
  {$ELSE}
    Classes,
    SysUtils,
    Windows,
    StdCtrls,
    ExtCtrls,
    {$IFDEF VER130}
      FileCtrl,
    {$ENDIF}
  {$ENDIF}
    QImport3ZipMcpt;

{ TBaseDocumentFile }

const
  LF = #13#10;
  sFileNameNotDefined = 'File name is not defined';
  sFileNotFound = 'File %s not found';

procedure TBaseDocumentFile.SetFileName(const Value: qiString);
begin
  if FFileName <> Value then;
  begin
    FFileName := Value;
    FLoaded := False;
  end;
end;

procedure TBaseDocumentFile.Decompress;

  function GetTempDir: string; 
  var
    Buffer: array[0..MAX_PATH] of AnsiChar;
  begin
    GetTempPathA(SizeOf(Buffer) - 1, Buffer);
    Result := String(StrPas(Buffer));
  end;

var
  ZipFile: TMiniZip;
  Guid: TGUID;
begin
  {$IFDEF VCL6}
  CreateGUID(Guid);
    FCurrFolder := IncludeTrailingPathDelimiter(GetTempDir) + GUIDToString(Guid) + '\';
  {$ELSE}
    FCurrFolder := ExtractFileDir(ParamStr(0)) + '\temp\';
  {$ENDIF}
  ZipFile := TMiniZip.Create(nil);
  try
    ZipFile.UnZipfile := AnsiString(FFileName);
    ZipFile.UnzipAllTo(FCurrFolder);
  finally
    ZipFile.Free;
  end;
end;

procedure TBaseDocumentFile.DeleteTempFolder;

  function FullRemoveDir(Dir: string; DeleteAllFilesAndFolders,
    StopIfNotAllDeleted, RemoveRoot: boolean): Boolean;
  var
    i: Integer;
    SRec: TSearchRec;
    FN: string;
  begin
{$IFDEF VCL6} 
  {$WARN SYMBOL_PLATFORM OFF}
{$ENDIF}
    Result := False;
    if not DirectoryExists(Dir) then
      exit;
    Result := True;
    Dir := IncludeTrailingBackslash(Dir);
    i := FindFirst(Dir + '*', faAnyFile, SRec);
    try
      while i = 0 do
      begin
        FN := Dir + SRec.Name;
        if (SRec.Attr = faDirectory) or (SRec.Attr = faDirectory + faArchive) then
        begin
          if (SRec.Name <> '') and (SRec.Name <> '.') and (SRec.Name <> '..') then
          begin
            if DeleteAllFilesAndFolders then
              FileSetAttr(FN, faArchive);
            Result := FullRemoveDir(FN, DeleteAllFilesAndFolders,
              StopIfNotAllDeleted, True);
            if not Result and StopIfNotAllDeleted then
              exit;
          end;
        end
        else
        begin
          if DeleteAllFilesAndFolders then
            FileSetAttr(FN, faArchive);
          Result := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.DeleteFile(FN);
          if not Result and StopIfNotAllDeleted then
            exit;
        end;
        i := FindNext(SRec);
      end;
    finally
      {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.FindClose(SRec);
    end;
    if not Result then
      exit;
    if RemoveRoot then
      if not RemoveDir(Dir) then
        Result := false;

{$IFDEF VCL6} 
  {$WARN SYMBOL_PLATFORM ON}
{$ENDIF}
  end;

begin
  FullRemoveDir(FCurrFolder, True, False, True);
end;

constructor TBaseDocumentFile.Create;
begin
  FLoaded := False;
end;

procedure TBaseDocumentFile.Load;
begin
  if FFileName = EmptyStr then
    raise Exception.Create(sFileNameNotDefined);
  if not FileExists(FFileName) then
    raise Exception.CreateFmt(sFileNotFound, [FFileName]);
  if not FLoaded then
    try
      Decompress;
      if DirectoryExists(FCurrFolder) then
      begin
        LoadXML(FCurrFolder);
        FLoaded := True;
      end;
    finally
      DeleteTempFolder;
    end;
end;

procedure TBaseDocumentFile.Clear;
begin
  FFileName := '';
  FLoaded := False;
end;

end.
