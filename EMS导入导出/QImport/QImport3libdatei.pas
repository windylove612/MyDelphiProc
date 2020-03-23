unit QImport3libdatei;
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


function suerAr(SS:String;Ar:array of String): String;

implementation




function StrLongint(ia,ib : string; Anfang:Integer):Integer;
var
LenABuf,LenBBuf,CurPos:Integer;
function IsBufOk: Boolean;
var J:Integer;
begin
 result:=false;
 for J:=1 to LenBBuf-1 do
  if ia[CurPos+J]<>ib[J+1] then exit;
 result:=true;
end;

begin
REsult:=0;
if (ia='') or (ib='') then exit;
LenABuf:=Length(ia);
if Anfang < 1 then
  Anfang:=1;
CurPos:=Anfang;
if CurPos>LenABuf then exit;

LenBBuf:=Length(ib);
dec(LenABuf,LenBBuf-2);
while CurPos < LenABuf do begin
   if (ia[CurPos]=ib[1]) and ((LenBBuf=1) or IsBufOk) then begin
     result:=CurPos; exit end;
   inc(CurPos);
 end;
end;


function SuErPos(st,su,er:String;WPos : DWord) : String;
var
h:longint;
begin
result:=st;
if wpos = 0 then exit;
h:= strlongint(st,su,WPos);
if h > 0 then
   begin   {:'#$D'#   :F}
   delete(st,h,length(su));
   if er<>'' then
       insert(er,st,h);
   end;
result:=st;
end;

function suerAr(SS:String;Ar:array of String): String;
var J,Max,x,LenErs:Integer;
begin
J:=Low(Ar); Max:=High(ar);
try
while J < Max do begin
  x:=pos(Ar[J],SS);
  while (x > 0) do begin
     SS:=suerpos(SS,Ar[J],Ar[J+1],x);
     LenErs:=length(Ar[J+1]);
     if LenErs=0 then
      LenErs:=1;
     x:=Strlongint(SS,Ar[J],x+LenErs);  //ohne +1 12.00 besser bei <br><br>
    end;
  inc(J,2);
 end; //while J
except
//interninfo('ex..suerAr');
end;
result:=SS;
end;

end.

