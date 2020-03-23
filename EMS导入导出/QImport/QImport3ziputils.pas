Unit QImport3ziputils;
{$I QImport3VerCtrl.Inc}
{ ziputils.pas - IO on .zip files using zlib
  - definitions, declarations and routines used by both
    zip.pas and unzip.pas
    The file IO is implemented here.

  based on work by Gilles Vollant

  March 23th, 2000,
  Copyright (C) 2000 Jacques Nomssi Nzali }

interface

//{$undef UseStream}
{$define Delphi}
{$ifdef UseStream}
  {$define Streams}
{$endif}

uses
  {$IFDEF Delphi}
    {$IFDEF VCL16}
      System.Classes,
      System.SysUtils,
      Winapi.Windows,
    {$ELSE}
      Classes,
      SysUtils,
      Windows,
    {$ENDIF}
  {$ENDIF}
  QImport3zutil;

{ -------------------------------------------------------------- }
{$ifdef Streams}
type
  FILEptr = TFileStream;
{$else}
type
  FILEptr = ^file;
{$endif}
type
  seek_mode = (SEEK_SET, SEEK_CUR, SEEK_END);
  open_mode = (fopenread, fopenwrite, fappendwrite);

function fopen(filename : PAnsiChar; mode : open_mode) : FILEptr;

procedure fclose(fp : FILEptr);

function fseek(fp : FILEptr; recPos : uLong; mode : seek_mode) : int_t;

function fread(buf : voidp; recSize : uInt;
               recCount : uInt; fp : FILEptr) : uInt;

function fwrite(buf : voidp;  recSize : uInt;
                recCount : uInt; fp : FILEptr) : uInt;

function ftell(fp : FILEptr) : uLong;  { ZIP }

function feof(fp : FILEptr) : uInt;   { MiniZIP }


{ ------------------------------------------------------------------- }

type
  zipFile = voidp;
  unzFile = voidp;
type
  z_off_t = long;

{ tm_zip contain date/time info }
type
  tm_zip = record
     tm_sec : uInt;            { seconds after the minute - [0,59] }
     tm_min : uInt;            { minutes after the hour - [0,59] }
     tm_hour : uInt;           { hours since midnight - [0,23] }
     tm_mday : uInt;           { day of the month - [1,31] }
     tm_mon : uInt;            { months since January - [0,11] }
     tm_year : uInt;           { years - [1980..2044] }
  end;

 tm_unz = tm_zip;

// "Added Stuff"
function Dofiletime(f : PAnsiChar;               { name of file to get info on }
   var tmzip : tm_zip; { return value: access, modific. and creation times }
   var dt : Int_t) : uLong;                { dostime }

procedure change_file_date(const filename : PAnsiChar;
                           dosdate : uLong;
                           tmu_date : tm_unz);

function tm_zipToFileTime(tmu_date : tm_unz):TFileTime;
 
const
  Z_BUFSIZE = (16384);
  Z_MAXFILENAMEINZIP = (256);

const
  CENTRALHEADERMAGIC = $02014b50;

const
  SIZECENTRALDIRITEM = $2e;
  SIZEZIPLOCALHEADER = $1e;

function ALLOC(size : int_t) : voidp;

procedure TRYFREE(p : voidp);

const
  Paszip_copyright : PAnsiChar = ' Paszip Copyright 2000 Jacques Nomssi Nzali ';

implementation

function ALLOC(size : int_t) : voidp;
begin
  ALLOC := zcalloc (NIL, size, 1);
end;

procedure TRYFREE(p : voidp);
begin
  if Assigned(p) then
    zcfree(NIL, p);
end;

{$ifdef Streams}
{ ---------------------------------------------------------------- }

function fopen(filename : PAnsiChar; mode : open_mode) : FILEptr;
var
  fp : FILEptr;
begin
  fp := NIL;
  try
    Case mode of
    fopenread: fp := TFileStream.Create(filename, fmOpenRead);
    fopenwrite: fp := TFileStream.Create(filename, fmCreate or fmOpenWrite);
    fappendwrite :
      begin
        fp := TFileStream.Create(filename, fmCreate or fmOpenReadWrite);
        fp.Seek(soFromEnd, 0);
      end;
    end;
  except
    on EFOpenError do
      fp := NIL;
  end;
  fopen := fp;
end;

procedure fclose(fp : FILEptr);
begin
  fp.Free;
end;

function fread(buf : voidp;
               recSize : uInt;
               recCount : uInt;
               fp : FILEptr) : uInt;
var
  totalSize, readcount : uInt;
begin
  if Assigned(buf) then
  begin
    totalSize := recCount * uInt(recSize);
    readCount := fp.Read(buf^, totalSize);
    if (readcount <> totalSize) then
      fread := readcount div recSize
    else
      fread := recCount;
  end
  else
    fread := 0;
end;

function fwrite(buf : voidp;
                recSize : uInt;
                recCount : uInt;
                fp : FILEptr) : uInt;
var
  totalSize, written : uInt;
begin
  if Assigned(buf) then
  begin
    totalSize := recCount * uInt(recSize);
    written := fp.Write(buf^, totalSize);
    if (written <> totalSize) then
      fwrite := written div recSize
    else
      fwrite := recCount;
  end
  else
    fwrite := 0;
end;

function fseek(fp : FILEptr;
               recPos : uLong;
               mode : seek_mode) : int;
const
  fsmode : array[seek_mode] of Word
    = (soFromBeginning, soFromCurrent, soFromEnd);
begin
  fp.Seek(recPos, fsmode[mode]);
  fseek := 0; { = 0 for success }
end;

function ftell(fp : FILEptr) : uLong;
begin
  ftell := fp.Position;
end;

function feof(fp : FILEptr) : uInt;
begin
  feof := 0;
  if Assigned(fp) then
    if fp.Position = fp.Size then
      feof := 1
    else
      feof := 0;
end;

{$else}
{ ---------------------------------------------------------------- }

function fopen(filename : PAnsiChar; mode : open_mode) : FILEptr;
var
  fp : FILEptr;
  OldFileMode : byte;
begin
  OldFileMode := FileMode;

  GetMem(fp, SizeOf(file));
  Assign(fp^, string(filename));
  {$i-}
  Case mode of
  fopenread:
    begin
      FileMode := 0;
      Reset(fp^, 1);
    end;
  fopenwrite:
    begin
      FileMode := 1;
      ReWrite(fp^, 1);
    end;
  fappendwrite :
    begin
      FileMode := 2;
      Reset(fp^, 1);
      Seek(fp^, FileSize(fp^));
    end;
  end;
  FileMode := OldFileMode;
  if IOresult<>0 then
  begin
    FreeMem(fp, SizeOf(file));
    fp := NIL;
  end;

  fopen := fp;
end;

procedure fclose(fp : FILEptr);
begin
  if Assigned(fp) then
  begin
    {$i-}
    system.close(fp^);
    if IOresult=0 then;
    FreeMem(fp, SizeOf(file));
  end;
end;

function fread(buf : voidp;
               recSize : uInt;
               recCount : uInt;
               fp : FILEptr) : uInt;
var
  totalSize, readcount : uInt;
begin
  if Assigned(buf) then
  begin
    totalSize := recCount * uInt(recSize);
    {$i-}
    system.BlockRead(fp^, buf^, totalSize, readcount);
    if (readcount <> totalSize) then
      fread := readcount div recSize
    else
      fread := recCount;
  end
  else
    fread := 0;
end;

function fwrite(buf : voidp;
                recSize : uInt;
                recCount : uInt;
                fp : FILEptr) : uInt;
var
  totalSize, written : uInt;
begin
  if Assigned(buf) then
  begin
    totalSize := recCount * uInt(recSize);
    {$i-}
    system.BlockWrite(fp^, buf^, totalSize, written);
    if (written <> totalSize) then
      fwrite := written div recSize
    else
      fwrite := recCount;
  end
  else
    fwrite := 0;
end;

function fseek(fp : FILEptr;
               recPos : uLong;
               mode : seek_mode) : int_t;
begin
  {$i-}
  case mode of
    SEEK_SET : system.Seek(fp^, recPos);
    SEEK_CUR : system.Seek(fp^, FilePos(fp^)+recPos);
    SEEK_END : system.Seek(fp^, FileSize(fp^)-1-recPos); { ?? check }
  end;
  fseek := IOresult; { = 0 for success }
end;

function ftell(fp : FILEptr) : uLong;
begin
  ftell := FilePos(fp^);
end;

function feof(fp : FILEptr) : uInt;
begin
  feof := 0;
  if Assigned(fp) then
    if eof(fp^) then
      feof := 1
    else
      feof := 0;
end;

{$endif}
{ ---------------------------------------------------------------- }


//# T.F. 03.02 .."Dostime" works till 2040 ...

function SystemTimeToZipTime(st:TSystemTime): tm_zip;
begin
 Result.tm_year := st.wYear;
 Result.tm_mon  := st.wMonth;
 Result.tm_mday := st.wDay;
 Result.tm_hour := st.wHour;
 Result.tm_min  := st.wMinute;
 Result.tm_sec  := st.wSecond;
end;

function UnZipTimeToSystemTime(tmu_date : tm_unz): TSystemTime;
begin
 Result.wYear:=tmu_date.tm_year;
 Result.wMonth:=tmu_date.tm_mon;
 Result.wDay:=tmu_date.tm_mday;
 Result.wHour:=tmu_date.tm_hour;
 Result.wMinute:=tmu_date.tm_min;
 Result.wSecond:=tmu_date.tm_sec;
 Result.wMilliseconds:=0;
end;


function tm_zipToFileTime(tmu_date : tm_unz): TFiletime;
begin
 SystemTimeToFileTime(UnZipTimeToSystemTime(tmu_date),Result);
end;

function Dofiletime(f : PAnsiChar;               { name of file to get info on }
   var tmzip : tm_zip; { return value: access, modifid. and creation times }
   var dt : Int_t) : uLong;                { dostime }
var
  ftLocal : TFileTime; // FILETIME;
  hFind : THandle; // HANDLE;
  ff32 : TWIN32FindData; //  WIN32_FIND_DATA;
  st:TSystemTime;
begin
  Result := 0;
  hFind := FindFirstFile( PChar(f), ff32);
  if (hFind <> INVALID_HANDLE_VALUE) then begin
    FileTimeToLocalFileTime(ff32.ftLastWriteTime,ftLocal);
    FileTimeToSystemTime(ftLocal,st);
    tmzip:=SystemTimeToZipTime(st);
    FileTimeToDosDateTime(ftLocal,LongRec(dt).hi,LongRec(dt).lo);
    FindClose(hFind);
    Result := 1;
  end;
end;



procedure change_file_date(const filename : PAnsiChar;
                           dosdate : uLong;
                           tmu_date : tm_unz);
var
  hFile : THandle;
  ftm,ftLocal,ftCreate,ftLastAcc,ftLastWrite : TFileTime;
  st: SystemTime;
begin
  hFile := CreateFile( PChar(filename),GENERIC_READ or GENERIC_WRITE,
                      0,NIL,OPEN_EXISTING,0,0);
  GetFileTime(hFile, @ftCreate, @ftLastAcc, @ftLastWrite);
  st:=UnZipTimeToSystemTime(tmu_date);
  SystemTimeToFileTime(st,ftLocal);
//  DosDateTimeToFileTime(LongRec(dosdate).hi,LongRec(dosdate).lo, ftLocal); //#T.F. o2.o3 WORD((dosdate shl 16)), WORD(dosdate) changed
  LocalFileTimeToFileTime(ftLocal, ftm);
  SetFileTime(hFile,@ftm, @ftLastAcc, @ftm);
  CloseHandle(hFile);
end;


end.
