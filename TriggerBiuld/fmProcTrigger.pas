unit fmProcTrigger;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, RzLabel, RzEdit, ExtCtrls, RzPanel, RzButton, Mask,comobj,
  DB, DBTables, Grids, DBGrids, RzDBGrid;

type                                     
  TForm1 = class(TForm)
    RzPanel1: TRzPanel;
    RzMemo1: TRzMemo;
    RzPanel2: TRzPanel;
    RzLabel1: TRzLabel;
    RzButton1: TRzButton;
    RzButton2: TRzButton;
    RzDBGrid1: TRzDBGrid;
    Query1: TQuery;
    tblLB: TTable;
    dsLB: TDataSource;
    tblLBTBLNAME: TStringField;
    OpenDialog1: TOpenDialog;
    RzLabel2: TRzLabel;
    edtUSER: TRzEdit;
    RzLabel3: TRzLabel;
    edtBM: TRzEdit;
    RzButton3: TRzButton;
    procedure RzButton1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure RzButton2Click(Sender: TObject);
    procedure RzButton3Click(Sender: TObject);
  private
    { Private declarations }
    Function ReadTextFromTextFile(sTextFile:String):String;
    Procedure GetFromTextFile(sLineText:String);
    function MakeTempFileName: string;
    procedure WriteLog(sLog: string);
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.GetFromTextFile(sLineText: String);
var
  iPos1,iPos2,i:integer;
  TmpStr,Str : String;
  sCD:string;
begin
  Str:=sLineText;
  tblLB.Append;
  try
    tblLBTBLNAME.AsString:=Str;
  except
    On E:Exception do
    begin
      ShowMessage('导入的文本文件格式有错：'+E.Message);
    end;
  end;
  tblLB.Post;
end;

function TForm1.MakeTempFileName: string;
var
  Hour, Min, Sec, MSec: Word;             //111
begin
  DecodeTime(Now, Hour, Min, Sec, MSec);
  Result := Format('_%.2d%.2d%.2d%.4d', [Hour, Min, Sec, MSec]);
end;

function TForm1.ReadTextFromTextFile(sTextFile: String): String;
var
  Str : String;
  i ,tmpBZ: integer;
  sList:TStringList;
  iPos1,iPos2:integer;
begin
  sList:=tstringlist.Create;
  try
    sList.LoadFromFile(PChar(sTextFile));
    For i:= 0 to sList.Count-1 do
    begin
      Str := sList.Strings[i];
      GetFromTextFile(Str);
    end;
  finally
    sList.Free;
  end;
end;

procedure TForm1.RzButton1Click(Sender: TObject);
var
  ExcelApp,MyWorkBook: Variant;
  Loop :integer;
  filename,sName :string;
begin
  OpenDialog1.FileName := '';
  OpenDialog1.Execute;
  if OpenDialog1.FileName = '' then
  begin
    exit;
  end;
  filename := OpenDialog1.FileName;
  tblLB.Close;
  tblLB.EmptyTable;
  tblLB.Open;
  ReadTextFromTextFile(filename);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  tblLB.TableName:=MakeTempFileName;
  tblLB.CreateTable;
  tblLB.Open;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  if tblLB.Exists then
  begin
    tblLB.close;
    tblLB.DeleteTable;
  end;
end;

procedure TForm1.RzButton2Click(Sender: TObject);
var
  sNR:string;
begin
  if tblLB.RecordCount<=0 then exit;
  RzMemo1.Lines.Clear;
  sNR:='';
  tblLB.First;
  while not tblLB.Eof do
  begin
    if edtUSER.Text<>'' then
    begin
      sNR:=sNR+'create sequence '+edtUSER.Text+'.'+'S_'+tblLBTBLNAME.AsString+';'+#13#10+'/'+#13#10;
      sNR:=sNR+'create or replace trigger '+edtUSER.Text+'.'+'TIB_'+tblLBTBLNAME.AsString
           +' before insert '+#13#10+'on '+edtUSER.Text+'.'+tblLBTBLNAME.AsString+' for each row'+#13#10
           +'declare'+#13#10
           +'     integrity_error  exception;'+#13#10
           +'     errno            integer;'+#13#10
           +'     errmsg           char(200);'+#13#10
           +'begin'+#13#10
           +'    select '+edtUSER.Text+'.'+'S_'+tblLBTBLNAME.AsString+'.NextVal into :NEW.TM from dual;'+#13#10
           +'exception'+#13#10
           +'    when integrity_error then'+#13#10
           +'         raise_application_error(errno, errmsg);'+#13#10
           +'end;'+#13#10
           +'/'+#13#10;
      sNR:=sNR+'create or replace trigger '+edtUSER.Text+'.'+'TIU_'+tblLBTBLNAME.AsString
           +' before update '+#13#10+'on '+edtUSER.Text+'.'+tblLBTBLNAME.AsString+' for each row'+#13#10
           +'declare'+#13#10
           +'     integrity_error  exception;'+#13#10
           +'     errno            integer;'+#13#10
           +'     errmsg           char(200);'+#13#10
           +'begin'+#13#10
           +'    select '+edtUSER.Text+'.'+'S_'+tblLBTBLNAME.AsString+'.NextVal into :NEW.TM from dual;'+#13#10
           +'exception'+#13#10
           +'    when integrity_error then'+#13#10
           +'         raise_application_error(errno, errmsg);'+#13#10
           +'end;'+#13#10
           +'/'+#13#10;
    end
    else
    begin
      sNR:=sNR+'create sequence '+'S_'+tblLBTBLNAME.AsString+';'+#13#10+'/'+#13#10;
      sNR:=sNR+'create or replace trigger '+'TIB_'+tblLBTBLNAME.AsString
           +' before insert '+#13#10+'on '+tblLBTBLNAME.AsString+' for each row'+#13#10
           +'declare'+#13#10
           +'     integrity_error  exception;'+#13#10
           +'     errno            integer;'+#13#10
           +'     errmsg           char(200);'+#13#10
           +'begin'+#13#10
           +'    select '+'S_'+tblLBTBLNAME.AsString+'.NextVal into :NEW.TM from dual;'+#13#10
           +'exception'+#13#10
           +'    when integrity_error then'+#13#10
           +'         raise_application_error(errno, errmsg);'+#13#10
           +'end;'+#13#10
           +'/'+#13#10;
      sNR:=sNR+'create or replace trigger '+'TIU_'+tblLBTBLNAME.AsString
           +' before update '+#13#10+'on '+tblLBTBLNAME.AsString+' for each row'+#13#10
           +'declare'+#13#10
           +'     integrity_error  exception;'+#13#10
           +'     errno            integer;'+#13#10
           +'     errmsg           char(200);'+#13#10
           +'begin'+#13#10
           +'    select '+'S_'+tblLBTBLNAME.AsString+'.NextVal into :NEW.TM from dual;'+#13#10
           +'exception'+#13#10
           +'    when integrity_error then'+#13#10
           +'         raise_application_error(errno, errmsg);'+#13#10
           +'end;'+#13#10
           +'/'+#13#10;
    end;
    RzMemo1.Lines.Add(sNR);
    tblLB.Next;
  end;
  if sNR<>'' then
    WriteLog(sNR)
end;

procedure TForm1.WriteLog(sLog: string);
var
  LogFile: TextFile;
  sFileName,str: string;
begin
  sFileName := '.\trigger.txt';
  try
    str := sLog;
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

procedure TForm1.RzButton3Click(Sender: TObject);
var
  sNR:string;
begin
  if edtBM.Text='' then exit;
  RzMemo1.Lines.Clear;
  sNR:='';
  if edtUSER.Text<>'' then
  begin
    sNR:=sNR+'create sequence '+edtUSER.Text+'.'+'S_'+edtBM.Text+';'+#13#10+'/'+#13#10;
    sNR:=sNR+'create or replace trigger '+edtUSER.Text+'.'+'TIB_'+edtBM.Text
         +' before insert '+#13#10+'on '+edtUSER.Text+'.'+edtBM.Text+' for each row'+#13#10
         +'declare'+#13#10
         +'     integrity_error  exception;'+#13#10
         +'     errno            integer;'+#13#10
         +'     errmsg           char(200);'+#13#10
         +'begin'+#13#10
         +'    select '+edtUSER.Text+'.'+'S_'+edtBM.Text+'.NextVal into :NEW.TM from dual;'+#13#10
         +'exception'+#13#10
         +'    when integrity_error then'+#13#10
         +'         raise_application_error(errno, errmsg);'+#13#10
         +'end;'+#13#10
         +'/'+#13#10;
    sNR:=sNR+'create or replace trigger '+edtUSER.Text+'.'+'TIU_'+edtBM.Text
         +' before update '+#13#10+'on '+edtUSER.Text+'.'+edtBM.Text+' for each row'+#13#10
         +'declare'+#13#10
         +'     integrity_error  exception;'+#13#10
         +'     errno            integer;'+#13#10
         +'     errmsg           char(200);'+#13#10
         +'begin'+#13#10
         +'    select '+edtUSER.Text+'.'+'S_'+edtBM.Text+'.NextVal into :NEW.TM from dual;'+#13#10
         +'exception'+#13#10
         +'    when integrity_error then'+#13#10
         +'         raise_application_error(errno, errmsg);'+#13#10
         +'end;'+#13#10
         +'/'+#13#10;
  end
  else
  begin
    sNR:=sNR+'create sequence '+'S_'+edtBM.Text+';'+#13#10+'/'+#13#10;
    sNR:=sNR+'create or replace trigger '+'TIB_'+edtBM.Text
         +' before insert '+#13#10+'on '+edtBM.Text+' for each row'+#13#10
         +'declare'+#13#10
         +'     integrity_error  exception;'+#13#10
         +'     errno            integer;'+#13#10
         +'     errmsg           char(200);'+#13#10
         +'begin'+#13#10
         +'    select '+'S_'+edtBM.Text+'.NextVal into :NEW.TM from dual;'+#13#10
         +'exception'+#13#10
         +'    when integrity_error then'+#13#10
         +'         raise_application_error(errno, errmsg);'+#13#10
         +'end;'+#13#10
         +'/'+#13#10;
    sNR:=sNR+'create or replace trigger '+'TIU_'+edtBM.Text
         +' before update '+#13#10+'on '+edtBM.Text+' for each row'+#13#10
         +'declare'+#13#10
         +'     integrity_error  exception;'+#13#10
         +'     errno            integer;'+#13#10
         +'     errmsg           char(200);'+#13#10
         +'begin'+#13#10
         +'    select '+'S_'+edtBM.Text+'.NextVal into :NEW.TM from dual;'+#13#10
         +'exception'+#13#10
         +'    when integrity_error then'+#13#10
         +'         raise_application_error(errno, errmsg);'+#13#10
         +'end;'+#13#10
         +'/'+#13#10;
  end;
  RzMemo1.Lines.Add(sNR);
  if sNR<>'' then
    WriteLog(sNR)
end;

end.
