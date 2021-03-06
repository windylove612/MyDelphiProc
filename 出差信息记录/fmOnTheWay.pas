unit fmOnTheWay;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RzButton, StdCtrls, RzEdit, RzCmboBx, Mask, RzLabel, ExtCtrls,
  RzPanel,IniFiles, RzRadChk, RzLstBox, RzChkLst, SQLSrch,ShellAPI;

type
  TForm1 = class(TForm)
    RzPanel1: TRzPanel;
    RzLabel1: TRzLabel;
    RQ: TRzDateTimeEdit;
    RzLabel2: TRzLabel;
    edtMC: TRzEdit;
    RzLabel4: TRzLabel;
    btnSure: TRzButton;
    btnAdd: TRzButton;
    RzButton1: TRzButton;
    RzButton2: TRzButton;
    RQ_END: TRzDateTimeEdit;
    RzLabel5: TRzLabel;
    edtTS: TRzEdit;
    RzLabel3: TRzLabel;
    edtDJ: TRzEdit;
    RzLabel6: TRzLabel;
    edtJK: TRzEdit;
    RzLabel7: TRzLabel;
    lstTX: TSQLCheckListBox;
    RzLabel8: TRzLabel;
    ckcBX: TRzCheckBox;
    RzButton3: TRzButton;
    RzMemo1: TRzMemo;
    RzButton4: TRzButton;
    RzButton5: TRzButton;
    OpenDialog1: TOpenDialog;
    RzButton6: TRzButton;
    RzButton7: TRzButton;
    procedure btnAddClick(Sender: TObject);
    procedure btnSureClick(Sender: TObject);
    procedure RzButton1Click(Sender: TObject);
    procedure RzButton2Click(Sender: TObject);
    procedure RzButton3Click(Sender: TObject);
    procedure RzButton4Click(Sender: TObject);
    procedure RzButton5Click(Sender: TObject);
    procedure RzButton6Click(Sender: TObject);
    procedure RzButton7Click(Sender: TObject);
  private
    { Private declarations }
    sSCDir:string;
    procedure Log(pStr: string;sLJ,sName:string);
    procedure ReLog(pStr: string;sLJ,sName:string);
    procedure WtiteLog(pStr: string);
    function GetCompactDate(pRQ: TDateTime): string;
    procedure GetNR();
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Log(pStr,sLJ,sName: string);
var
  dat: TextFile;
  sFile: string;
begin
  sFile := sName + '.txt';
  try
    if not DirectoryExists(sLJ) then
      MkDir(sLJ);
    Assignfile(dat, sLJ + '\' + sFile);
    if FileExists(sLJ + '\' + sFile) then
      Append(dat)
    else
      Rewrite(dat);
    Writeln(dat,  pStr);
    Flush(dat);
  finally
    Closefile(dat);
  end;
end;

procedure TForm1.WtiteLog(pStr: string);
var
  LogFile: TextFile;
  sFileName,str: string;
begin
  sFileName := GetCurrentDir+'\出差信息\'+GetCompactDate(RQ.Date)+'TO'+edtMC.Text+'.txt';
  try
    str := pStr;
    AssignFile(LogFile, sFileName);
    if FileExists(sFileName) then
      Rewrite(LogFile)
    else
      Rewrite(LogFile);
    Writeln(LogFile, str);
    Flush(LogFile);
    CloseFile(LogFile);
  finally
  end;
end;

function TForm1.GetCompactDate(pRQ: TDateTime): string;
begin
  Result := FormatDateTime('yyyymmdd', pRQ);
end;

procedure TForm1.GetNR;
var
  Ini: TIniFile;
  sSecList: TStringList;
  sSection:string;
  sRQ:string;
  sTX:string;
  sBX:string;
  i,j,k:integer;
  sTmp:string;
  sXH:integer;
begin
  ini := TIniFile.Create(GetCurrentDir+'\出差信息\'+GetCompactDate(RQ.Date)+'TO'+edtMC.Text+'.txt');
  sSection:=edtMC.Text;
  try
    sRQ := Ini.ReadString(sSection, '出差日期', '');
    i:=pos('-',sRQ);
    RQ.Date:=StrToDate(copy(sRQ,1,i-1));
    RQ_END.Date:=StrToDate(copy(sRQ,i+1,length(sRQ)));
    edtMC.Text := Ini.ReadString(sSection, '出差地点', '');
    edtTS.Text := Ini.ReadString(sSection, '出差天数', '');
    edtDJ.Text := Ini.ReadString(sSection, '房费单价', '');
    edtJK.Text := Ini.ReadString(sSection, '借款情况', '');
    sBX:= Ini.ReadString(sSection, '是否已报销', '');
    if sBX='是' then
      ckcBX.Checked:=true
    else
      ckcBX.Checked:=false;
    sTX:= Ini.ReadString(sSection, '周末可调休天数统计', '');
    lstTX.Items.Clear;
    lstTX.Values.Clear;
    sTmp:=sTX;
    for j:=0 to maxint do
    begin
      i:=pos(';',sTmp);
      if i=0 then
      begin
        if copy(sTmp,1,2)='NO' then
        begin
          k:=pos('NO',sTmp);
          if k>0 then
          begin
            lstTX.Items.add(copy(sTmp,k+3,length(sTmp)));
            lstTX.Values.add(copy(sTmp,k+3,length(sTmp)));
            sXH:=lstTX.Values.IndexOf(copy(sTmp,k+3,length(sTmp)));
            lstTX.ItemChecked[sXH]:=false;
          end;
        end;
        if copy(sTmp,1,3)='YES' then
        begin
          k:=pos('YES',sTmp);
          if k>0 then
          begin
            lstTX.Items.add(copy(sTmp,k+4,length(sTmp)));
            lstTX.Values.add(copy(sTmp,k+4,length(sTmp)));
            sXH:=lstTX.Values.IndexOf(copy(sTmp,k+4,length(sTmp)));
            lstTX.ItemChecked[sXH]:=true;
          end;
        end;
        break;
      end
      else
      begin
        if copy(sTmp,1,2)='NO' then
        begin
          k:=pos('NO',sTmp);
          if k>0 then
          begin
            lstTX.Items.add(copy(sTmp,k+3,i-4));
            lstTX.Values.add(copy(sTmp,k+3,i-4));
            sXH:=lstTX.Values.IndexOf(copy(sTmp,k+3,i-4));
            lstTX.ItemChecked[sXH]:=false;
          end;
        end;
        if copy(sTmp,1,3)='YES' then
        begin
          k:=pos('YES',sTmp);
          if k>0 then
          begin
            lstTX.Items.add(copy(sTmp,k+4,i-5));
            lstTX.Values.add(copy(sTmp,k+4,i-5));
            sXH:=lstTX.Values.IndexOf(copy(sTmp,k+4,i-5));
            lstTX.ItemChecked[sXH]:=true;
          end;
        end;
        sTmp:=copy(sTmp,i+1,length(sTmp));
      end;
    end;
  finally
    Ini.Free;
  end;
end;

procedure TForm1.ReLog(pStr, sLJ, sName: string);
var
  dat: TextFile;
  sFile: string;
begin
  sFile := sName + '.txt';
  try
    if not DirectoryExists(sLJ) then
      MkDir(sLJ);
    Assignfile(dat, sLJ + '\' + sFile);
    if FileExists(sLJ + '\' + sFile) then
      Rewrite(dat)
    else
      Rewrite(dat);
    Writeln(dat,  pStr);
    Flush(dat);
  finally
    Closefile(dat);
  end;
end;

procedure TForm1.btnAddClick(Sender: TObject);
var
  sRQ:Tdatetime;
  i:integer;
begin
  if (RQ.Text='') or (RQ_END.Text='') then
  begin
    showmessage('日期录入不完整');
    exit;
  end;
  edtTS.Text:=IntToStr(trunc(RQ_END.Date)-trunc(RQ.Date)+1);
  //判断是不是周末
  lstTX.Items.Clear;
  lstTX.Values.Clear;
  sRQ:=RQ.date;
  for i:=0 to StrToInt(edtTS.Text)-1 do
  begin
    if (Integer(DayOfWeek(sRQ+i))=1) or (Integer(DayOfWeek(sRQ+i))=7) then
    begin
      lstTX.Items.add(DateToStr(sRQ+i));
      lstTX.Values.add(DateToStr(sRQ+i));
    end;
  end;
end;

procedure TForm1.btnSureClick(Sender: TObject);
var
  sNR:string;
  i:integer;
  sTX:string;
begin
  if (RQ.Text='') or (RQ_END.Text='') then
  begin
    showmessage('日期录入不完整');
    exit;
  end;
  if edtMC.Text='' then
  begin
    showmessage('请录入出差地点！');
    exit;
  end;
  if edtTS.Text='' then
  begin
    showmessage('请录入出差天数！');
    exit;
  end;
  if edtDJ.Text='' then
  begin
    showmessage('请录入房费单价！');
    exit;
  end;
  if edtJK.Text='' then
  begin
    showmessage('请录入借款情况！');
    exit;
  end;
  sNR:='';
  sNR:=sNR+'['+edtMC.Text+']'+#13#10;
  sNR:=sNR+'出差日期='+DateToStr(RQ.date)+'-'+DateToStr(RQ_END.date)+#13#10;
  sNR:=sNR+'出差地点='+edtMC.Text+#13#10;
  sNR:=sNR+'出差天数='+edtTS.Text+#13#10;
  sNR:=sNR+'房费单价='+edtDJ.Text+#13#10;
  sNR:=sNR+'借款情况='+edtJK.Text+#13#10;
  sTX:='';
  for i:=0 to lstTX.Count-1 do
  begin
    if lstTX.ItemChecked[i] then
      sTX:=sTX+';'+'YES-'+lstTX.Values[i]
    else
      sTX:=sTX+';'+'NO-'+lstTX.Values[i];
  end;
  sTX:=copy(sTX,2,length(sTX));
  sNR:=sNR+'周末可调休天数统计='+sTX+#13#10;
  if ckcBX.Checked then
    sNR:=sNR+'是否已报销=是'+#13#10
  else
    sNR:=sNR+'是否已报销=否'+#13#10;
  WtiteLog(sNR);
  showmessage('录入完毕');
end;

procedure TForm1.RzButton1Click(Sender: TObject);
begin
  if RQ.Text='' then
  begin
    showmessage('日期录入不完整');
    exit;
  end;
  if edtMC.Text='' then
  begin
    showmessage('请录入出差地点！');
    exit;
  end;
  if FileExists(GetCurrentDir+'\出差信息\'+GetCompactDate(RQ.Date)+'TO'+edtMC.Text+'.txt') then
    GetNR;
end;

procedure TForm1.RzButton2Click(Sender: TObject);
begin
  RQ.Clear;
  RQ_END.Clear;
  edtMC.Text:='';
  edtTS.Text:='';
  edtDJ.Text:='';
  edtJK.Text:='';
  lstTX.Items.Clear;
  lstTX.Values.Clear;
  ckcBX.Checked:=false;
end;

procedure TForm1.RzButton3Click(Sender: TObject);
var
  sList:TStringList;
  i:integer;
  Str:string;
begin
  if FileExists(GetCurrentDir+'\出差模版\出差申请.txt') then
  begin
    sList:=tstringlist.Create;
    try
      RzMemo1.Lines.Clear;
      sList.LoadFromFile(PChar(GetCurrentDir+'\出差模版\出差申请.txt'));
      For i:= 0 to sList.Count-1 do
      begin
        Str := sList.Strings[i];
        RzMemo1.Lines.add(Str);
      end;
    finally
      sList.Free;
    end;
  end;
end;

procedure TForm1.RzButton4Click(Sender: TObject);
var
  sList:TStringList;
  i:integer;
  Str:string;
begin
  if FileExists(GetCurrentDir+'\出差模版\借款申请.txt') then
  begin
    sList:=tstringlist.Create;
    try
      RzMemo1.Lines.Clear;
      sList.LoadFromFile(PChar(GetCurrentDir+'\出差模版\借款申请.txt'));
      For i:= 0 to sList.Count-1 do
      begin
        Str := sList.Strings[i];
        RzMemo1.Lines.add(Str);
      end;
    finally
      sList.Free;
    end;
  end;
end;

procedure TForm1.RzButton5Click(Sender: TObject);
begin
  ShellExecute(Handle, 'open', PChar('explorer.exe'), PChar(GetCurrentDir+'\出差模版'), nil, SW_SHOW);
end;

procedure TForm1.RzButton6Click(Sender: TObject);
begin
  ShellExecute(Handle, 'open', PChar('calc.exe'), nil, nil, SW_SHOW);
  //WinExec('calc.exe',SW_SHOWNORMAL);
end;

procedure TForm1.RzButton7Click(Sender: TObject);
begin
  ShellExecute(Handle, 'open', PChar('explorer.exe'), PChar(GetCurrentDir+'\出差信息'), nil, SW_SHOW);
end;

end.
