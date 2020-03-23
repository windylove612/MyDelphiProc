unit fmProcRecord;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, RzButton, StdCtrls, RzEdit, RzCmboBx, Mask, RzLabel, ExtCtrls,
  RzPanel,IniFiles,ShellAPI, OleServer, OutlookXP;

type
  TForm1 = class(TForm)
    RzPanel1: TRzPanel;
    RzLabel1: TRzLabel;
    RQ: TRzDateTimeEdit;
    RzLabel2: TRzLabel;
    edtMC: TRzEdit;
    RzLabel3: TRzLabel;
    ComXM: TRzComboBox;
    edtNR: TRzMemo;
    RzLabel4: TRzLabel;
    btnSure: TRzButton;
    btnAdd: TRzButton;
    RzButton1: TRzButton;
    RzButton2: TRzButton;
    FontDialog1: TFontDialog;
    RzButton3: TRzButton;
    btnShow: TRzButton;
    RzButton4: TRzButton;
    RzButton5: TRzButton;
    RzButton6: TRzButton;
    RzButton7: TRzButton;
    RzButton8: TRzButton;
    RzButton9: TRzButton;
    RzButton10: TRzButton;
    procedure FormCreate(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnSureClick(Sender: TObject);
    procedure ComXMChange(Sender: TObject);
    procedure RzButton2Click(Sender: TObject);
    procedure RzButton1Click(Sender: TObject);
    procedure RzButton3Click(Sender: TObject);
    procedure RzButton4Click(Sender: TObject);
    procedure RzButton5Click(Sender: TObject);
    procedure RzButton6Click(Sender: TObject);
    procedure RzButton7Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure RQKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure ComXMKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edtMCKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure edtNRKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure RzButton8Click(Sender: TObject);
    procedure RzButton9Click(Sender: TObject);
    procedure RzButton10Click(Sender: TObject);
  private
    { Private declarations }
    sSCDir:string;
    procedure DoFillList();
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

procedure TForm1.DoFillList;
var
  sList:TStringList;
  i:integer;
  Str:string;
begin
  if FileExists(GetCurrentDir+'.\ProcList.txt') then
  begin
    sList:=tstringlist.Create;
    try
      ComXM.Items.Clear;
      ComXM.Values.Clear;
      ComXM.Items.add(' ');
      ComXM.Values.add(' ');
      sList.LoadFromFile(PChar(GetCurrentDir+'.\ProcList.txt'));
      For i:= 0 to sList.Count-1 do
      begin
        Str := sList.Strings[i];
        ComXM.Items.add(Str);
        ComXM.Values.add(Str);
      end;
    finally
      sList.Free;
    end;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  DoFillList;
  if not DirectoryExists(GetCurrentDir+'\项目需求修改') then
    MkDir(GetCurrentDir+'\项目需求修改');
end;

procedure TForm1.btnAddClick(Sender: TObject);
begin
  if edtMC.Text='' then
  begin
    showmessage('请输入项目名称！');
    exit;
  end;
  if ComXM.Values.IndexOf(edtMC.Text)>=0 then
  begin
    showmessage('该项目名称已存在');
    exit;
  end;
  WtiteLog(edtMC.Text);
  showmessage('添加完毕');
  DoFillList;
end;

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
  sFileName := '.\ProcList.txt';
  try
    str := pStr;
    AssignFile(LogFile, sFileName);
    if FileExists(sFileName) then
      Append(LogFile)
    else
      Rewrite(LogFile);
    Writeln(LogFile, str);
    Flush(LogFile);
    CloseFile(LogFile);
  finally
  end;
end;

procedure TForm1.btnSureClick(Sender: TObject);
begin
  if RQ.Text='' then
  begin
    showmessage('请选择日期');
    exit;
  end;
  if edtMC.Text='' then
  begin
    showmessage('请输入项目名称！');
    exit;
  end;
  if edtNR.Lines.Text='' then
  begin
    showmessage('请输入内容');
    exit;
  end;
  Log(edtNR.Lines.Text,GetCurrentDir+'\项目需求修改\'+GetCompactDate(RQ.Date),edtMC.Text);
  showmessage('录入完成');
end;

procedure TForm1.ComXMChange(Sender: TObject);
begin
  edtMC.Text:=ComXM.Values[ComXM.itemIndex];
end;

function TForm1.GetCompactDate(pRQ: TDateTime): string;
begin
  Result := FormatDateTime('yyyymmdd', pRQ);
end;

procedure TForm1.RzButton2Click(Sender: TObject);
begin
  edtNR.Lines.Clear;
end;

procedure TForm1.RzButton1Click(Sender: TObject);
begin
  if RQ.Text='' then
  begin
    showmessage('请选择日期');
    exit;
  end;
  if edtMC.Text='' then
  begin
    showmessage('请输入项目名称');
    exit;
  end;
  if FileExists(GetCurrentDir+'\项目需求修改\'+GetCompactDate(RQ.Date)+'\'+edtMC.Text+'.txt') then
    GetNR
  else
    edtNR.Lines.Clear;
end;

procedure TForm1.GetNR;
var
  sList:TStringList;
  i:integer;
  Str:string;
begin
  sList:=tstringlist.Create;
  try
    edtNR.Lines.Clear;
    sList.LoadFromFile(PChar(GetCurrentDir+'\项目需求修改\'+GetCompactDate(RQ.Date)+'\'+edtMC.Text+'.txt'));
    For i:= 0 to sList.Count-1 do
    begin
      Str := sList.Strings[i];
      edtNR.Lines.add(Str);
    end;
  finally
    sList.Free;
  end;
end;

procedure TForm1.RzButton3Click(Sender: TObject);
begin
  if FontDialog1.Execute then
  begin
    btnShow.Font := FontDialog1.Font;
    edtNR.Font := FontDialog1.Font;
  end;
end;

procedure TForm1.RzButton4Click(Sender: TObject);
begin
  if RQ.Text='' then
  begin
    showmessage('请选择日期');
    exit;
  end;
  if edtMC.Text='' then
  begin
    showmessage('请输入项目名称！');
    exit;
  end;
  if edtNR.Lines.Text='' then
  begin
    showmessage('请输入内容');
    exit;
  end;
  ReLog(edtNR.Lines.Text,GetCurrentDir+'\项目需求修改\'+GetCompactDate(RQ.Date),edtMC.Text);
  showmessage('录入完成');
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

procedure TForm1.RzButton5Click(Sender: TObject);
begin
  ShellExecute(Handle, 'open', PChar('E:\生活实用\压缩\ASPack.exe'), nil, nil, SW_SHOW);
  //WinExec('ASPack.exe',SW_SHOWNORMAL);
end;

procedure TForm1.RzButton6Click(Sender: TObject);
begin
  WinExec('sqlmon.exe',SW_SHOWNORMAL);
end;

procedure TForm1.RzButton7Click(Sender: TObject);
begin
  ShellExecute(Handle, 'open', PChar('explorer.exe'), PChar(GetCurrentDir+'\项目需求修改'), nil, SW_SHOW);
end;

procedure TForm1.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=vk_f1 then
    btnSureClick(btnSure);
  if key=vk_f2 then
    RzButton4Click(RzButton4);
  if key=vk_f3 then
    RzButton2Click(RzButton2);
  if key=vk_f4 then
    RzButton5Click(RzButton5);
  if key=vk_f5 then
    RzButton6Click(RzButton6);
  if key=vk_f6 then
    RzButton7Click(RzButton7);
  if key=vk_f7 then
    RzButton1Click(RzButton1);
  if key=vk_f8 then
    RzButton8Click(RzButton8);
  if key=vk_f9 then
    RzButton9Click(RzButton9);
  if key=vk_f10 then
    RzButton10Click(RzButton10);
  if key=vk_f12 then
  begin
    showmessage('DeSigner By gaoxiang');
  end;
end;

procedure TForm1.RQKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=vk_f1 then
    btnSureClick(btnSure);
  if key=vk_f2 then
    RzButton4Click(RzButton4);
  if key=vk_f3 then
    RzButton2Click(RzButton2);
  if key=vk_f4 then
    RzButton5Click(RzButton5);
  if key=vk_f5 then
    RzButton6Click(RzButton6);
  if key=vk_f6 then
    RzButton7Click(RzButton7);
  if key=vk_f7 then
    RzButton1Click(RzButton1);
  if key=vk_f8 then
    RzButton8Click(RzButton8);
  if key=vk_f9 then
    RzButton9Click(RzButton9);
  if key=vk_f10 then
    RzButton10Click(RzButton10);
  if key=vk_f12 then
  begin
    showmessage('DeSigner By gaoxiang');
  end;
end;

procedure TForm1.ComXMKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=vk_f1 then
    btnSureClick(btnSure);
  if key=vk_f2 then
    RzButton4Click(RzButton4);
  if key=vk_f3 then
    RzButton2Click(RzButton2);
  if key=vk_f4 then
    RzButton5Click(RzButton5);
  if key=vk_f5 then
    RzButton6Click(RzButton6);
  if key=vk_f6 then
    RzButton7Click(RzButton7);
  if key=vk_f7 then
    RzButton1Click(RzButton1);
  if key=vk_f8 then
    RzButton8Click(RzButton8);
  if key=vk_f9 then
    RzButton9Click(RzButton9);
  if key=vk_f10 then
    RzButton10Click(RzButton10);
  if key=vk_f12 then
  begin
    showmessage('DeSigner By gaoxiang');
  end;
end;

procedure TForm1.edtMCKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=vk_f1 then
    btnSureClick(btnSure);
  if key=vk_f2 then
    RzButton4Click(RzButton4);
  if key=vk_f3 then
    RzButton2Click(RzButton2);
  if key=vk_f4 then
    RzButton5Click(RzButton5);
  if key=vk_f5 then
    RzButton6Click(RzButton6);
  if key=vk_f6 then
    RzButton7Click(RzButton7);
  if key=vk_f7 then
    RzButton1Click(RzButton1);
  if key=vk_f8 then
    RzButton8Click(RzButton8);
  if key=vk_f9 then
    RzButton9Click(RzButton9);
  if key=vk_f10 then
    RzButton10Click(RzButton10);
  if key=vk_f12 then
  begin
    showmessage('DeSigner By gaoxiang');
  end;
end;

procedure TForm1.edtNRKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key=vk_f1 then
    btnSureClick(btnSure);
  if key=vk_f2 then
    RzButton4Click(RzButton4);
  if key=vk_f3 then
    RzButton2Click(RzButton2);
  if key=vk_f4 then
    RzButton5Click(RzButton5);
  if key=vk_f5 then
    RzButton6Click(RzButton6);
  if key=vk_f6 then
    RzButton7Click(RzButton7);
  if key=vk_f7 then
    RzButton1Click(RzButton1);
  if key=vk_f8 then
    RzButton8Click(RzButton8);
  if key=vk_f9 then
    RzButton9Click(RzButton9);
  if key=vk_f10 then
    RzButton10Click(RzButton10);
  if key=vk_f12 then
  begin
    showmessage('DeSigner By gaoxiang');
  end;
end;

procedure TForm1.RzButton8Click(Sender: TObject);
begin
  ShellExecute(Handle, 'open', PChar(GetCurrentDir+'\Tools\Tran_CDNR.exe'), nil, nil, SW_SHOW);
end;

procedure TForm1.RzButton9Click(Sender: TObject);
begin
  ShellExecute(Handle, 'open', PChar(GetCurrentDir+'\Tools\Pre_JM_CDNR.exe'), nil, nil, SW_SHOW);
end;

procedure TForm1.RzButton10Click(Sender: TObject);
begin
  ShellExecute(Handle, 'open', PChar(GetCurrentDir+'\Tools\ProcPswd.exe'), nil, nil, SW_SHOW);
end;

end.
