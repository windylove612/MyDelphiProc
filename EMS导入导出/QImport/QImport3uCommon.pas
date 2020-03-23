unit QImport3uCommon;
{$I QImport3VerCtrl.Inc}
interface
uses
  {$IFDEF VCL16}
    System.SysUtils;
  {$ELSE}
    SysUtils;
  {$ENDIF}


function makeDirs(nurPfad :string): boolean;
function isdir(sos:string):boolean;
function Rightdir(adir :string): string;

implementation


function Rightdir(adir :string): string;
var
hs : string;

begin
result:=adir;
If adir='' then exit;
hs:=adir;     //ansiuppercase ??
//while spos(hs,' ') = 1 do hs:=suer(hs,' ','');
if hs[length(hs)]='\' then
   result:=hs else
     result:=hs + '\';
end;


function Rightdir2(adir :string): string;
begin
result:=adir;
If adir='' then exit;
if result[length(result)]='\' then
     delete(result,length(result),1);
end;

function isdir(sos:string):boolean;
var
sr: TSearchRec;
L : Integer;

begin
result:=false;
If sos='' then exit;
sos:=rightdir2(sos);
L:=length(sos);
if (l=0) or (sos[L]='.') or (sos[1]='.') then exit;
result:=(findfirst( sos,faanyfile,sr) = 0) and (sr.attr=sr.attr or fadirectory);
FindClose(sr);
end;

function snurpfad(ss :string):string;
var ts:string;
begin  result:='';
if length(ss) < 3 then exit;
ts:=extractfilepath(ss);
result:=copy(ts,1,length(ts)-1);
end;


function makeDirs(nurPfad :string): boolean;
const maxVersuche=10; //je subdir
label 1,2,3;
var
i,ir : integer;
Pnurpfad,mkpfad : string;

begin
result:=false;
If (length(nurpfad) < 4) then
 exit;
try
Pnurpfad:=nurPfad;
if Pnurpfad[length(Pnurpfad)]='\' then
delete(Pnurpfad,length(Pnurpfad),1);

result:=isdir(pnurpfad);

if result  then
   exit;

{$I-}   // $I+  weiter unten
i:=0;
1:    mkpfad:=Pnurpfad;
      mkDir(mkpfad);   ir:=IOResult;
        if  (ir <> 0) and (ir <> 5)then begin
2:      if mkpfad = '' then
           exit;
        mkDir(mkpfad);  inc(i);
        if i= maxVersuche then begin
            exit //i:=0; if not BInfo(mkpfad,'Verzeichnisse weiter erstellen? (makeDirs)',mb_yesno) then exit;
           end;
        if  IOResult <> 0 then begin
             mkpfad:=snurpfad(mkpfad);
             i:=0; //08.00
             goto 2
           end;// else goto 3;
        goto 1;
        end;
3:
result :=true;
except
result:=false; end;
end;


end.

