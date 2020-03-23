unit CSRQ;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,unConnectDB, DB, DBTables, StdCtrls, ComCtrls, Grids, DBGrids,
  RzDBGrid, RzButton;

type
  TfmProcCSRQ_SFZ = class(TForm)
    GroupBox1: TGroupBox;
    Memo1: TMemo;
    LabMsg: TLabel;
    pb1: TProgressBar;
    Button1: TButton;
    Database1: TDatabase;
    tmpQry: TQuery;
    qrySave: TQuery;
    DBGrid1: TRzDBGrid;
    tblSFZ: TTable;
    dsSFZ: TDataSource;
    Label1: TLabel;
    RzButton1: TRzButton;
    SaveDialog: TSaveDialog;
    tblSFZSFZBH: TStringField;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure dsSFZDataChange(Sender: TObject; Field: TField);
    procedure RzButton1Click(Sender: TObject);
  private
    { Private declarations }
    function IsValidSFZ(psSFZ: string): Boolean;
    function OldSFZToNew(pOLD: string): string;
    procedure WriteDebugLog(sLog: string);
  public
    { Public declarations }
  end;

var
  fmProcCSRQ_SFZ: TfmProcCSRQ_SFZ;

implementation

{$R *.dfm}

procedure TfmProcCSRQ_SFZ.FormCreate(Sender: TObject);
begin
  ConnectDB(Database1,'DB.ini');
  tblSFZ.TableName:='xx';
  tblSFZ.CreateTable;
  tblSFZ.Open;
end;

procedure TfmProcCSRQ_SFZ.Button1Click(Sender: TObject);
var
  iYear, iMonth, iDay,i,j,flag,iXB: Integer;
  dateCSRQ: TDatetime;
  s:string;
begin
  flag:=0;
  with tmpQry do
  begin
    SQL.Text := Memo1.Text;
    Open();
    pb1.Max := RecordCount;
    pb1.Position := 0;
    while not eof do
    begin
     { //�ж����֤�Ƿ���ȷ
      for i:=0 to Length(tmpQry.FieldByName('SFZBH').AsString)-1 do
      begin
        s:=copy(tmpQry.FieldByName('SFZBH').AsString,i+1,1);
        for j:=0 to 9 do
        begin
          if ((s=IntToStr(j)) or (s='X') or (s='x')) then
          begin
            flag:=1;
            break;
          end;
        end;
        if flag=0 then
        begin
          showmessage('�����֤��'+tmpQry.FieldByName('SFZBH').AsString+'����ȷ������');
          if not tblSFZ.Locate('SFZBH',tmpQry.FieldByName('SFZBH').AsString,[]) then
          begin
            tblSFZ.Append;
            tblSFZSFZBH.AsString:=tmpQry.FieldByName('SFZBH').AsString;
            tblSFZ.Post;
          end;
          break;
        end;
      end;
      if flag=0 then
      begin
        Next();
        continue;
      end;   }

      if Length(trim(tmpQry.FieldByName('SFZBH').AsString)) = 18 then                     //���֤��Ϊ18λ��
      begin
        iYear := StrToInt(Copy(tmpQry.FieldByName('SFZBH').AsString, 7, 4));
        iMonth := StrToInt(Copy(tmpQry.FieldByName('SFZBH').AsString, 11, 2));
        iDay := StrToInt(Copy(tmpQry.FieldByName('SFZBH').AsString, 13, 2));
        iXB := StrToInt(Copy(tmpQry.FieldByName('SFZBH').AsString, 17, 1));
      end
      else if Length(trim(tmpQry.FieldByName('SFZBH').AsString)) = 15 then                //���֤��Ϊ15λ��
      begin
        iYear := 1900 + StrToInt(Copy(tmpQry.FieldByName('SFZBH').AsString, 7, 2));
        iMonth := StrToInt(Copy(tmpQry.FieldByName('SFZBH').AsString, 9, 2));
        iDay := StrToInt(Copy(tmpQry.FieldByName('SFZBH').AsString, 11, 2));
        iXB := StrToInt(Copy(tmpQry.FieldByName('SFZBH').AsString, 15, 1));
      end
      else
      begin
        if Application.MessageBox(PChar('���֤��'+tmpQry.FieldByName('SFZBH').AsString+'����15��18λ���Ƿ����,�������úţ�'),'ȷ��',MB_YESNO)=ID_NO then
        begin
          if not tblSFZ.Locate('SFZBH',tmpQry.FieldByName('SFZBH').AsString,[]) then
          begin
            tblSFZ.Append;
            tblSFZSFZBH.AsString:=tmpQry.FieldByName('SFZBH').AsString;
            tblSFZ.Post;
          end;
          Next();
          continue;
        end;
      end;
      if (iMonth > 12) or (iMonth < 1) or (iDay > 31) or (iDay < 1) or ((iMonth = 2) and (iDay >= 30)) then
      begin
        if Application.MessageBox(PChar('���֤��'+tmpQry.FieldByName('SFZBH').AsString+'���ն������Ƿ����,�������úţ�'),'ȷ��',MB_YESNO)=ID_NO then
        begin
          if not tblSFZ.Locate('SFZBH',tmpQry.FieldByName('SFZBH').AsString,[]) then
          begin
            tblSFZ.Append;
            tblSFZSFZBH.AsString:=tmpQry.FieldByName('SFZBH').AsString;
            tblSFZ.Post;
          end;
          Next();
          continue;
        end
        else
        begin
          iYear := StrToInt(Copy(tmpQry.FieldByName('SFZBH').AsString, 7, 4));
          iMonth := StrToInt(Copy(tmpQry.FieldByName('SFZBH').AsString, 11, 2));
          iDay := StrToInt(Copy(tmpQry.FieldByName('SFZBH').AsString, 13, 2));
          iXB := StrToInt(Copy(tmpQry.FieldByName('SFZBH').AsString, 17, 1));
        end;
      end;
     { if not IsValidSFZ(tmpQry.FieldByName('SFZBH').AsString) then
      begin
        if Application.MessageBox(PChar('���֤��'+tmpQry.FieldByName('SFZBH').AsString+'��֤���󣬸����֤����Ϊ�Ƿ����룬�Ƿ����,�������úţ�'),'ȷ��',MB_YESNO)=ID_NO then
        begin
          if not tblSFZ.Locate('SFZBH',tmpQry.FieldByName('SFZBH').AsString,[]) then
          begin
            tblSFZ.Append;
            tblSFZSFZBH.AsString:=tmpQry.FieldByName('SFZBH').AsString;
            tblSFZ.Post;
          end;
          Next();
          continue;
        end;
      end; }
      dateCSRQ := Encodedate(iYear, iMonth, iDay);
      with qrySave do
      begin
        SQL.Text:='begin';
        SQL.Add('update HYK_GRXX set CSRQ=:CSRQ,SEX=:SEX where HYID=:HYID;');
        sql.Add('end;');
        ParamByName('CSRQ').AsDateTime:=dateCSRQ;
        if (iXB mod 2)=1 then
          parambyname('SEX').AsInteger:= 0
        else if (iXB mod 2)=0 then
          parambyname('SEX').AsInteger:= 1;
        parambyname('HYID').AsInteger:=tmpQry.FieldByName('HYID').AsInteger;
        ExecSQL;
      end;
      Repaint;
      Next();
      pb1.Position := pb1.Position + 1;
      LabMsg.Caption := Format('%d/%d', [pb1.Position, pb1.Max]);
      LabMsg.Update;
    end;
    Close();
  end;
end;

function TfmProcCSRQ_SFZ.IsValidSFZ(psSFZ: string): Boolean;
begin
  Result := psSFZ = OldSFZToNew(Copy(psSFZ, 1, 6) + Copy(psSFZ, 9, 9));
end;

function TfmProcCSRQ_SFZ.OldSFZToNew(pOLD: string): string;
const
  W: array[1..18] of integer = (7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2, 1);
  A: array[0..10] of char = ('1', '0', 'x', '9', '8', '7', '6', '5', '4', '3', '2');
var
  i, j, S: integer;
begin
  if Length(pOLD) <> 15 then
    Result := ''
  else
  begin
    Result := pOLD;
    Insert('19', Result, 7);
    S := 0;
    try
      for i := 1 to 17 do
      begin
        j := StrToInt(Result[i]) * W[i];
        S := S + j;
      end;
    except
      result := '';
      exit;
    end;
    S := S mod 11;
    Result := Result + A[S];
  end;
end;

procedure TfmProcCSRQ_SFZ.FormDestroy(Sender: TObject);
begin
  tblSFZ.close;
  tblSFZ.DeleteTable;
end;

procedure TfmProcCSRQ_SFZ.dsSFZDataChange(Sender: TObject; Field: TField);
begin
  Label1.Caption:=IntToStr(tblSFZ.RecordCount);
end;

procedure TfmProcCSRQ_SFZ.RzButton1Click(Sender: TObject);
var
  TxtFile: TextFile;
  str: string;
begin
  tblSFZ.First;
  while not tblSFZ.Eof do
  begin
    str:=tblSFZSFZBH.AsString;
    WriteDebugLog(str);
    tblSFZ.Next;
  end;
end;

procedure TfmProcCSRQ_SFZ.WriteDebugLog(sLog: string);
var
  LogFile: TextFile;
  sFileName,str: string;
begin
  sFileName := '.\SFZBH_ERROR.log';
  try
    str := sLog;
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

end.
