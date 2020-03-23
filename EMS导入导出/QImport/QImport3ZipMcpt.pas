unit QImport3ZipMcpt;
{$I QImport3VerCtrl.Inc}
interface

uses
  {$IFDEF VCL16}
    System.Classes,
    System.SysUtils,
    Winapi.Windows,
  {$ELSE}
    Classes,
    SysUtils,
    Windows,
  {$ENDIF}
  QImport3ucommon,
  QImport3zlib110,
  QImport3zutil,
  QImport3libdatei,
  QImport3ziputils,
  QImport3zip,
  QImport3unzip;

type
  PrgsType = (ptZip, ptUnzip, ptZipStart, ptZipping);
  //TProgressEvent used by Eric Engler "ZipMaster" Zip/Unzip
  TPrgsEvent   = procedure(Sender: TObject; ProgrType: PrgsType; Filename: AnsiString;
                                                        uncompressed_size: Integer) of object;
  TListFileEvent   = procedure(Sender: TObject; LastWriteTime : TFileTime; Filename: AnsiString;
                                       Ratio, uncompressed_size, compressed_size:Integer) of object;
  TUnzipFileEvent   = procedure(Sender: TObject; LastWriteTime : TFileTime;
                          uncompressed_size, compressed_size: Integer;
                           DestFilename: AnsiString; var CanOverwrite: Boolean) of object;

                        //zip_fileinfo
    TMiniZip = class(TComponent)
    private
      FzipFile : zipFile;
      FUnZFile : unzFile;
      FOnProgress : TPrgsEvent;
      FOnUnzipFile : TUnzipFileEvent;
      FOnListFile: TListFileEvent;
      SZipfile, SUnZipfile : AnsiString;
      FUzWithoutPath,FUzOverwrite, BGetSize :Boolean;
      buf : pointer;
      FUnzAllComprSize : Integer;
     procedure OpenZipfile;
     procedure OpenUnZipfile;
     procedure SetUnZipfile(Value:AnsiString);
     procedure SetZipfile(Value:AnsiString);
     function  GetUnzAllComprSize: integer;
     function  do_extract_currentfile(DestDir:string) : integer;
    public
     MemStream : TMemoryStream;
     constructor Create(AOwner: TComponent); override;
     Destructor  Destroy; override;
     procedure OpenAppendZipfile;  //Use it only to add to a existing Zipfile
     function ListUnzFiles: Boolean;
     function AddToZipFile(FileName,FileName_InZip:AnsiString): Integer;
     function UnzExtractFileToMemStream(FileNameInZip:AnsiString) : boolean;
     function UnzExtractFileTo(FileNameInZip,DestDir:AnsiString) : boolean;
     function UnzipAllTo(DestDir:string): Boolean;
     property Zipfile : AnsiString Read SZipfile write SetZipfile;
     property UnZipfile : AnsiString Read SUnZipfile write SetUnZipfile;
     property UnzAllUnComprSize : Integer read GetUnzAllComprSize;
    published
     property OnProgress:    TPrgsEvent   read  FOnProgress  write FOnProgress;
     property OnListFile:    TListFileEvent   read  FOnListFile  write FOnListFile;
     property OnUnzipFile:   TUnzipFileEvent  read  FOnUnzipFile write FOnUnzipFile;
     property UzWithoutPath : Boolean read FUzWithoutPath write FUzWithoutPath;
     property UzOverwrite : Boolean read FUzOverwrite write FUzOverwrite;
end;

procedure Register;

implementation

//{$R *.dcr}
const
   WRITEBUFFERSIZE = 8192;

procedure Register;
begin
  RegisterComponents('FreeWare', [TMiniZip]);
end;

constructor TMiniZip.Create(AOwner: TComponent);
begin
inherited;
FUzWithoutPath:=false;        FUnzAllComprSize:=-1;
FUzOverwrite:=true;           BGetSize:=false;
if csdesigning in componentstate then exit;
buf := zcalloc (NIL, WRITEBUFFERSIZE, 1);
end;

Destructor  TMiniZip.Destroy;
begin
inherited;
if csdesigning in componentstate then exit;
try
zcfree(NIL, buf);
if (FUnZFile<>nil) then begin
  unzClose(FUnZFile); FUnZFile:=Nil; end;
if (FZipFile<>nil) then begin
  zipClose(FZipFile,nil); FZipFile:=Nil; end;
if assigned(MemStream) then
  MemStream.Free;
except end;  
end;


function TMiniZip.GetUnzAllComprSize: Integer;
begin
Result:=FUnzAllComprSize;
if FUnzAllComprSize >-1 then exit;

BGetSize:=true;
if ListUnzFiles then
   Result:=FUnzAllComprSize;
BGetSize:=false;
end;

procedure TMiniZip.OpenZipfile;
begin
if (FZipFile<>nil) then exit;  //or not fileexists(SZipfile)
FZipFile:=zipOpen(PAnsiChar(SZipfile),0);
end;

procedure TMiniZip.OpenAppendZipfile;
begin
if (FZipFile<>nil) then
     ZipClose(FZipFile,nil);
FZipFile:=zipOpen(PAnsiChar(SZipfile),1);
end;
procedure TMiniZip.OpenUnZipfile;
begin
if (FUnZFile<>nil) or not fileexists(string(SUnZipfile)) then exit;
FUnZFile:=unzOpen(PAnsiChar(SUnZipfile));
end;

procedure TMiniZip.SetUnZipfile(Value:AnsiString);
begin
if SUnZipfile=Value then exit;
if (FUnZFile<>nil) then begin
  unzClose(FUnZFile); FUnZFile:=Nil; end;
SUnZipfile:=Value;
end;

procedure TMiniZip.SetZipfile(Value:AnsiString);
begin
if SZipfile=Value then exit;
if (FZipFile<>nil) then begin
  ZipClose(FZipFile,nil); FZipFile:=Nil; end;
SZipfile:=Value;
end;

Function TMiniZip.ListUnzFiles: Boolean;
var
  i : int_t;
  gi : unz_global_info;
  filename_inzip : array[0..255] of AnsiChar;
  file_info : unz_file_info;
  ratio : uLong;
begin
Result:=false;
OpenUnZipfile;
if (not BGetSize and not assigned(FOnListFile)) or (FUnZFile=Nil) then exit;

if unzGetGlobalInfo(FUnZFile, gi) <> UNZ_OK then exit;
unzGoToFirstFile(FUnZFile);

//  WriteLn(' Length  Method   Size  Ratio   Date    Time   CRC-32     Name');
FUnzAllComprSize:=0;
for i := 0 to gi.number_entry-1 do begin
    ratio := 0;
    if unzGetCurrentFileInfo(FUnZFile, @file_info, filename_inzip,
       sizeof(filename_inzip),NIL,0,NIL,0) <> UNZ_OK then exit;

    if (file_info.uncompressed_size>0) then
      ratio := (file_info.compressed_size*100) div file_info.uncompressed_size;
    if not BGetSize then
    FOnListFile(Self,tm_zipToFileTime(file_info.tmu_date),AnsiString(filename_inzip),
           Ratio, file_info.uncompressed_size,file_info.compressed_size);

    FUnzAllComprSize:=FUnzAllComprSize+file_info.uncompressed_size;
    if ((i+1)<gi.number_entry) then begin
      if unzGoToNextFile(FUnZFile) <> UNZ_OK then exit;
    end;
end; //for

Result:=true;
end;


function TMiniZip.AddToZipFile(FileName,FileName_InZip:AnsiString): int_t;
var size_read :Integer;
    fin : FILEptr;
    zi  : zip_fileinfo;
begin
OpenZipfile;
if not fileexists(string(FileName)) or (FileName_InZip='') or (FzipFile=Nil) then begin
 Result:=ZIP_OK-1; exit end;


fillchar(zi,sizeof(zip_fileinfo),0);
Dofiletime(PAnsiChar(FileName),zi.tmz_date,zi.dosDate);

Result := zipOpenNewFileInZip(FzipFile,PAnsiChar(FileName_InZip), @zi,
                   NIL,0,NIL,0,NIL, Z_DEFLATED, Z_BEST_COMPRESSION);
                                      //Z_BEST_COMPRESSION
if (Result <> ZIP_OK) then exit;
 //  showmessage('Cannot OpenNewFileInZip'); exit end;

Result:=ZIP_OK-1;
if (buf=Nil) then exit;

fin := fopen(PAnsiChar(FileName), fopenread);
if (fin=NIL) then exit;
  // showmessage('error in opening  '+ESomeFile.Text); exit end;

if assigned(FOnProgress) then begin
 FOnProgress(self,ptZipStart,FileName,fseek(fin,0,SEEK_END));
 fseek(fin,0,SEEK_Set);
end;
 
 repeat
  size_read := fread(buf,1,WRITEBUFFERSIZE,fin);
  if assigned(FOnProgress) then
     FOnProgress(self,ptZipping,FileName,size_read);

  Result := ZIP_OK;
  if (size_read < WRITEBUFFERSIZE) and (feof(fin)=0) then exit;

  if (size_read>0) then try
    Result := zipWriteInFileInZip (FzipFile,buf,size_read);
  if (Result<0) then begin
    Result := ZIP_OK-1;exit; end;
  except Result := ZIP_OK-1; exit end;
 until (Result <> ZIP_OK) or (size_read=0);

 fclose(fin);
 if (Result<0) then
    Result := ZIP_ERRNO else begin
      Result := zipCloseFileInZip(FzipFile);
      if (Result<>ZIP_OK) then exit;
     end;
end;

function TMiniZip.UnzExtractFileToMemStream(FileNameInZip:AnsiString) : boolean;
begin
Result:=false;  
OpenUnZipfile;
if unzLocateFile(FUnZFile,PAnsiChar(FileNameInZip),2) <> UNZ_OK then exit;
if MemStream=Nil then
 MemStream := TMemoryStream.Create;
MemStream.Clear;
Result:=do_extract_currentfile('WriteToMemStream')=UNZ_OK;
MemStream.Position:=0;
end;

function TMiniZip.UnzExtractFileTo(FileNameInZip,DestDir:AnsiString) : boolean;
begin
Result:=false;
OpenUnZipfile;
if unzLocateFile(FUnZFile,PAnsiChar(FileNameInZip),2) <> UNZ_OK then exit;
Result:=do_extract_currentfile(string(DestDir))=UNZ_OK;
end;

function TMiniZip.do_extract_currentfile(DestDir:string) : int_t;
var
  filename_inzip : packed array[0..255] of AnsiChar;
  fout : FILEptr;
  file_info : unz_file_info;
  write_filename : AnsiString;
  ftestexist : FILEptr;
  CanOverwrite,UseMemStream : Boolean;
begin
  fout := NIL;

  Result := unzGetCurrentFileInfo(FUnZFile, @file_info, filename_inzip,
    sizeof(filename_inzip), NIL, 0, NIL,0);

  if (Result <> UNZ_OK) then begin
   // WriteLn('error ',Result, ' with zipfile in unzGetCurrentFileInfo');
    exit;
  end;
 UseMemStream:=DestDir='WriteToMemStream';
 write_filename := AnsiString(filename_inzip);
 if FUzWithoutPath then
   write_filename := AnsiString(extractfilename(string(write_filename)));

  Result := unzOpenCurrentFile(FUnZFile);
  if (Result <> UNZ_OK) then exit;
     // WriteLn('error ',Result,' with zipfile in unzOpenCurrentFile');


 write_filename:=AnsiString(suerar(DestDir+string(write_filename),['/','\']));
 if not UseMemStream and not makedirs(extractfilepath(string(write_filename))) then exit;

 CanOverwrite:=FUzOverwrite;
 if assigned(FOnUnzipFile) then
  FOnUnzipFile(Self,tm_zipToFileTime(file_info.tmu_date),file_info.Uncompressed_size,
   file_info.compressed_size,write_filename,CanOverwrite);

 if not UseMemStream and not CanOverwrite and (Result=UNZ_OK) then begin
      ftestexist := fopen(PAnsiChar(write_filename),fopenread);
      if (ftestexist <> NIL) then begin
        fclose(ftestexist);
        Result:= UNZ_OK+1;
        exit end;
    end;

 if not UseMemStream then begin
 fout := fopen(PAnsiChar(write_filename),fopenwrite);
 if (fout = NIL) then exit; end;

 repeat
  Result := unzReadCurrentFile(FUnZFile,buf,WRITEBUFFERSIZE);
  if assigned(FOnProgress) then
    FOnProgress(self,ptUnzip,write_filename,Result);
  if (Result<0) then exit;
     //WriteLn('error ',Result,' with zipfile in unzReadCurrentFile');
  if (Result>0) then begin
   if UseMemStream then
     MemStream.Write(buf^,Result) else
   if (fwrite(buf,Result,1,fout) <> 1) then begin
          // WriteLn('error in writing extracted file');
      Result := UNZ_ERRNO;  break;
     end;
   end;

 until (Result=0);

fclose(fout);
 if (Result=0) then
        change_file_date(PAnsiChar(write_filename),file_info.dosDate,
	   file_info.tmu_date);

if (Result=UNZ_OK) then begin
   Result := unzCloseCurrentFile(FUnZFile);
   if (Result <> UNZ_OK) then exit
       //	WriteLn('error ',Result,' with zipfile in unzCloseCurrentFile')
      else
        unzCloseCurrentFile(FUnZFile); { don't lose the error }
  end;
end;  //do_extract_currentfile


function TMiniZip.UnzipAllTo(DestDir:string) : Boolean;

var
  i : int_t;
  gi : unz_global_info;
  err : int_t;
begin
Result:=false;    OpenUnZipfile;
  err := unzGetGlobalInfo (FUnZFile, gi);
  if (err <> UNZ_OK) then exit;
    //WriteLn('error ',err,' with zipfile in unzGetGlobalInfo ');
DestDir:=RightDir(DestDir);
unzGoToFirstFile(FUnZFile);

  for i:=0 to gi.number_entry-1 do begin
    if (do_extract_currentfile(DestDir) <> UNZ_OK) then
      break;

    if ((i+1)<gi.number_entry) then begin
      err := unzGoToNextFile(FUnZFile);
      if (err <> UNZ_OK) then exit;
    end;
  end;

Result:=true;
end;


end.
