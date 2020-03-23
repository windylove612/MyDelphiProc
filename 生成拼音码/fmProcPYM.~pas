unit fmProcPYM;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DBTables, DB, ComCtrls, RzButton, StdCtrls, Mask, RzEdit,
  RzLabel, ExtCtrls, RzPanel,unConnectDB;

type
  TForm1 = class(TForm)
    RzPanel1: TRzPanel;
    Database1: TDatabase;
    RzLabel1: TRzLabel;
    RzLabel2: TRzLabel;
    RzLabel3: TRzLabel;
    RzLabel4: TRzLabel;
    edtTBL: TRzEdit;
    edtKEY: TRzEdit;
    edtZD: TRzEdit;
    btnOK: TRzButton;
    btnExit: TRzButton;
    LabMsg: TLabel;
    pb1: TProgressBar;
    tblPYM: TTable;
    qryTmp: TQuery;
    RzLabel5: TRzLabel;
    edtZH: TRzEdit;
    RzLabel6: TRzLabel;
    edtTJ: TRzEdit;
    RzLabel7: TRzLabel;
    RzLabel8: TRzLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
  private
    { Private declarations }
    function FoundFirstLetter(TempString: string): string;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  ConnectDB(Database1,'DB.ini');
  tblPYM.TableName:='xx';
  //tblPYM.CreateTable;
  //tblPYM.Open;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  tblPYM.close;
  tblPYM.DeleteTable;
end;

procedure TForm1.btnOKClick(Sender: TObject);
var
  i,j:integer;
  sSL:integer;
  sZD:string;
  sKEY:string;
  sZJ:string;
  sPYM:string;
  sZDX:string;
begin
  sSL:=1;
  if edtTBL.Text='' then
  begin
    showmessage('请输入更新表名');
    exit;
  end;
  if edtKEY.Text='' then
  begin
    showmessage('请输入主键名');
    exit;
  end;
  if edtZH.Text='' then
  begin
    showmessage('请输入转换字段名');
    exit;
  end;
  if edtZD.Text='' then
  begin
    showmessage('请输入更新字段名');
    exit;
  end;
  sZD:=edtZD.Text;
  sKEY:=edtKEY.Text;
  for i:=0 to Maxint do
  begin
    j:=pos(',',sZD);
    if j>0 then
    begin
      sZD:=copy(sZD,j+1,length(sZD));
      sSL:=sSL+1;
    end
    else
      break;
  end;
  sZD:=edtZD.Text;

  with qryTmp do
  begin
    close;
    sql.Clear;
    sql.Text:='select '+edtKEY.Text+','+edtZH.Text+' from '+edtTBL.Text+' where 1=1';
    if edtTJ.Text<>'' then
      sql.add('and '+edtTJ.Text);
    sql.add('order by '+edtKEY.Text);
    open;
  end;
  tblPYM.close;
  tblPYM.BatchMove(qryTmp,batcopy);
  tblPYM.Open;

  Database1.StartTransaction();
  try
    pb1.Max := tblPYM.RecordCount;
    pb1.Position := 0;
    tblPYM.First;
    while not tblPYM.Eof do
    begin
      sPYM:='';
      sZD:=edtZD.Text;
      sKEY:=edtKEY.Text;
      for i:=1 to sSL do
      begin
        j:=pos(',',sZD);
        if j>0 then
        begin
          sZDX:=copy(sZD,1,j-1);
          sPYM := sPYM + ',' + sZDX+'='''+FoundFirstLetter(tblPYM.FieldByName(edtZH.Text).AsString)+'''';
          sZD:=copy(sZD,j+1,length(sZD));
        end
        else
        begin
          sPYM:=sPYM+','+sZD+'='''+FoundFirstLetter(tblPYM.FieldByName(edtZH.Text).AsString)+'''';
          break;
        end;
      end;
      with qryTmp do
      begin
        close;
        sql.Clear;
        sql.Text:='update '+edtTBL.Text+' set '+copy(sPYM,2,length(sPYM))+' where 1=1';
        for i:=0 to maxint do
        begin
          j:=pos(',',sKEY);
          if j>0 then
          begin
            sZJ:=copy(sKEY,1,j-1);
            sql.add('and '+sZJ+'='+tblPYM.FieldByName(sZJ).AsString);
            sKEY:=copy(sKEY,j+1,length(sKEY));
          end
          else
          begin
            sql.add('and '+sKEY+'='''+tblPYM.FieldByName(sKEY).AsString+'''');
            break;
          end;
        end;
        execsql;
      end;
      tblPYM.Next;
      pb1.Position := pb1.Position + 1;
      LabMsg.Caption := Format('%d/%d', [pb1.Position, pb1.Max]);
      LabMsg.Update;
    end;
    Database1.Commit;
  except
    on E:Exception do
    begin
      showmessage(tblPYM.FieldByName(sKEY).AsString+tblPYM.FieldByName(sZJ).AsString);
      Database1.Rollback;
      ApplicationHandleException(Self);
      exit;
    end;
  end;
end;

function TForm1.FoundFirstLetter(TempString: string): string;
var
  List: array[1..26] of string;
  Last: array[1..26] of integer;
  TempChar: string;
  AscValue: integer;
  Long: integer;
  i: integer;
  Code: integer;
  ResultString: string;
begin
  ResultString := '';
  List[1] := 'a';
  Last[1] := 1637;
  List[2] := 'b';
  Last[2] := 1832;
  List[3] := 'c';
  Last[3] := 2077;
  List[4] := 'd';
  Last[4] := 2273;
  List[5] := 'e';
  Last[5] := 2301;
  List[6] := 'f';
  Last[6] := 2432;
  List[7] := 'g';
  Last[7] := 2593;
  List[8] := 'h';
  Last[8] := 2786;
  List[9] := 'i';
  Last[9] := 2786;
  List[10] := 'j';
  Last[10] := 3105;
  List[11] := 'k';
  Last[11] := 3211;
  List[12] := 'l';
  Last[12] := 3471;
  List[13] := 'm';
  Last[13] := 3634;
  List[14] := 'n';
  Last[14] := 3721;
  List[15] := 'o';
  Last[15] := 3729;
  List[16] := 'p';
  Last[16] := 3857;
  List[17] := 'q';
  Last[17] := 4026;
  List[18] := 'r';
  Last[18] := 4085;
  List[19] := 's';
  Last[19] := 4389;
  List[20] := 't';
  Last[20] := 4557;
  List[21] := 'u';
  Last[21] := 4557;
  List[22] := 'v';
  Last[22] := 4557;
  List[23] := 'w';
  Last[23] := 4683;
  List[24] := 'x';
  Last[24] := 4924;
  List[25] := 'y';
  Last[25] := 5248;
  List[26] := 'z';
  Last[26] := 5589;
  while length(TempString) <> 0 do
  begin
    TempChar := Copy(TempString, 1, 1);
    AscValue := Ord(TempChar[1]);
    if AscValue > 127 then
    begin
      TempChar := Copy(TempString, 1, 2);
      AscValue := Ord(TempChar[2]);
      if AscValue > 127 then
      begin
        Code := 100 * (Ord(TempChar[1]) - 160) + Ord(TempChar[2]) - 160;
        for i := 1 to 26 do
        begin
          if Code < Last[i] then
          begin
            ResultString := ResultString + UpperCase(List[i]);
            Break;
          end;
        end;
      end
      else
        ResultString := ResultString + TempChar[2];
      Long := Length(TempString);
      TempString := Copy(TempString, 3, Long - 2);
    end
    else
    begin
      Long := Length(TempString);
      TempString := Copy(TempString, 2, Long - 1);
      ResultString := ResultString + TempChar;
    end;
  end;
  if length(ResultString)>6 then
    ResultString:=copy(ResultString,1,6);
  StringReplace(ResultString,#39,'/',[rfReplaceAll]);
  Result := ResultString;
  if ResultString='' then
    Result := ' ';
end;

procedure TForm1.btnExitClick(Sender: TObject);
begin
  close;
end;

end.
