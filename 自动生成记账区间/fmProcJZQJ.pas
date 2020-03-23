unit fmProcJZQJ;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, RzEdit, RzLabel, RzButton, DB, DBTables,
  RzPanel, RzRadGrp, ExtCtrls,unConnectDB;

type
  TForm1 = class(TForm)
    RzPanel1: TRzPanel;
    rgLX: TRzRadioGroup;
    Database1: TDatabase;
    qryTmp: TQuery;
    RzButton1: TRzButton;
    RzLabel1: TRzLabel;
    edtBM: TRzEdit;
    RzLabel3: TRzLabel;
    KSRQ: TRzEdit;
    RzLabel4: TRzLabel;
    JSRQ: TRzEdit;
    rgJZ: TRzRadioGroup;
    RzButton2: TRzButton;
    RzButton3: TRzButton;
    qryConnect: TQuery;
    lblZY: TRzLabel;
    procedure RzButton1Click(Sender: TObject);
    procedure rgJZClick(Sender: TObject);
    procedure RzButton3Click(Sender: TObject);
    procedure RzButton2Click(Sender: TObject);
    procedure rgLXClick(Sender: TObject);
  private
    { Private declarations }
    function FullLeftStr(sMP:string;Len:Integer;FillStr:Char):string;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.RzButton1Click(Sender: TObject);
begin
  if rgLX.ItemIndex<0 then
  begin
    showmessage('请选择数据库类型');
    exit;
  end;
  if rgLX.ItemIndex=0 then
    ConnectDB(Database1,'DB_SYS.ini')
  else
    ConnectDB(Database1,'DB_ORA.ini');
  lblZY.Caption:='连接数据库成功';
  showmessage('连接数据库成功');
end;

procedure TForm1.rgJZClick(Sender: TObject);
begin
  if rgJZ.ItemIndex=0 then
  begin
    if rgLX.ItemIndex=1 then
    begin
      edtBM.Text:='CR_JZQJ';
    end
    else
    begin
      edtBM.Text:='JZQJ';
    end;
  end;
  if rgJZ.ItemIndex=1 then
  begin
    if rgLX.ItemIndex=1 then
    begin
      edtBM.Text:='CR_JZQJ_JD';
    end
    else
    begin
      edtBM.Text:='JZQJ_JD';
    end;
  end;
end;

procedure TForm1.RzButton3Click(Sender: TObject);
begin
  close;
end;

procedure TForm1.RzButton2Click(Sender: TObject);
var
  sY1,sY2:integer;
  i,j:integer;
  sRQ:TdateTime;
  sTS:integer;
  sDAY:integer;
begin
  if KSRQ.Text='' then
  begin
    showmessage('请输入开始年！');
    exit;
  end;
  if JSRQ.Text='' then
  begin
    showmessage('请输入结束年！');
    exit;
  end;
  sY1:=StrToInt(KSRQ.Text);
  sY2:=StrToInt(JSRQ.Text);

  if rgJZ.ItemIndex=0 then
  begin
    qryConnect.Close;
    qryConnect.Open;
    Database1.StartTransaction;
    try
      for i:=sY1 to sY2 do
      begin
        for j:=1 to 12 do
        begin
          if j=2 then
          begin
            sRQ:=EncodeDate(i,3,1)-1;
            with qryTmp do
            begin
              close;
              sql.Clear;
              sql.Text:='select to_number(to_char( :RQ ,''DD'')) from dual';
              ParamByName('RQ').AsDateTime:=sRQ;
              open;
              sTS:=fields[0].AsInteger;
            end;
          end;
          if ((j=1) or (j=3) or (j=5) or (j=7) or (j=8) or (j=10) or (j=12)) then
          begin
            sDAY:=31;
          end;
          if ((j=4) or (j=6) or (j=9) or (j=11)) then
          begin
            sDAY:=30;
          end;
          if j=2 then
          begin
            sDAY:=sTS;
          end;
          with qryTmp do
          begin
            close;
            sql.Clear;
            sql.Text:='insert into '+edtBM.text+'(YEARMONTH,KSRQ,JSRQ)';
            sql.Add('values(:YEARMONTH,:KSRQ,:JSRQ)');
            ParamByName('YEARMONTH').AsInteger:= StrToInt(IntToStr(i)+FullLeftStr(IntToStr(j),2,'0'));
            ParamByName('KSRQ').AsDateTime:=EncodeDate(i,j,1);
            ParamByName('JSRQ').AsDateTime:=EncodeDate(i,j,sDAY);
            execsql;
          end;
        end;
      end;
      Database1.Commit;
      qryConnect.Close;
      showmessage('生成成功');
    except
      Database1.Rollback;
      qryConnect.Close;
      ApplicationHandleException(Self);
      exit;
    end;
  end
  else
  begin
    qryConnect.Close;
    qryConnect.Open;
    Database1.StartTransaction;
    try
      for i:=sY1 to sY2 do
      begin
        for j:=1 to 4 do
        begin
          with qryTmp do
          begin
            close;
            sql.Clear;
            sql.Text:='insert into '+edtBM.text+'(JD,KSRQ,JSRQ,KSNY,JSNY)';
            sql.Add('values(:JD,:KSRQ,:JSRQ,:KSNY,:JSNY)');
            ParamByName('JD').AsInteger:=StrToInt(IntToStr(i)+IntToStr(j));
            if j=1 then
            begin
              ParamByName('KSRQ').AsDateTime:=EncodeDate(i,1,1);
              ParamByName('JSRQ').AsDateTime:=EncodeDate(i,3,31);
            end;
            if j=2 then
            begin
              ParamByName('KSRQ').AsDateTime:=EncodeDate(i,4,1);
              ParamByName('JSRQ').AsDateTime:=EncodeDate(i,6,30);
            end;
            if j=3 then
            begin
              ParamByName('KSRQ').AsDateTime:=EncodeDate(i,7,1);
              ParamByName('JSRQ').AsDateTime:=EncodeDate(i,9,30);
            end;
            if j=4 then
            begin
              ParamByName('KSRQ').AsDateTime:=EncodeDate(i,10,1);
              ParamByName('JSRQ').AsDateTime:=EncodeDate(i,12,31);
            end;
            if j=1 then
            begin
              ParamByName('KSNY').AsInteger:=StrToInt(IntToStr(i)+'01');
              ParamByName('JSNY').AsInteger:=StrToInt(IntToStr(i)+'03');
            end;
            if j=2 then
            begin
              ParamByName('KSNY').AsInteger:=StrToInt(IntToStr(i)+'04');
              ParamByName('JSNY').AsInteger:=StrToInt(IntToStr(i)+'06');
            end;
            if j=3 then
            begin
              ParamByName('KSNY').AsInteger:=StrToInt(IntToStr(i)+'07');
              ParamByName('JSNY').AsInteger:=StrToInt(IntToStr(i)+'09');
            end;
            if j=4 then
            begin
              ParamByName('KSNY').AsInteger:=StrToInt(IntToStr(i)+'10');
              ParamByName('JSNY').AsInteger:=StrToInt(IntToStr(i)+'12');
            end;
            execsql;
          end;
        end;
      end;
      Database1.Commit;
      qryConnect.Close;
      showmessage('生成成功');
    except
      Database1.Rollback;
      qryConnect.Close;
      ApplicationHandleException(Self);
      exit;
    end;
  end;
end;

procedure TForm1.rgLXClick(Sender: TObject);
begin
  rgJZClick(rgJZ);
end;

function TForm1.FullLeftStr(sMP:string; Len: Integer; FillStr: Char): string;
begin
  Result:= StringOfChar(FillStr, Len - Length(sMP)) + sMP;
end;

end.
