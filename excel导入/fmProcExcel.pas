unit fmProcExcel;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, DBTables, Grids, DBGrids, RzDBGrid, RzButton, RzCommon,
  StdCtrls, Mask, RzEdit, RzLabel,unConnectDB,comobj;

type
  TForm1 = class(TForm)
    RzLabel1: TRzLabel;
    edtTBL: TRzEdit;
    RzFrameController: TRzFrameController;
    BtnSJDR: TRzButton;
    RzDBGrid1: TRzDBGrid;
    qryExcel: TQuery;
    tblExcel: TTable;
    dsExcel: TDataSource;
    Database1: TDatabase;
    OpenDialog1: TOpenDialog;
    RzButton1: TRzButton;
    tblExcelHYK_NO: TStringField;
    qryConnect: TQuery;
    qryID: TQuery;
    procedure BtnSJDRClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure RzButton1Click(Sender: TObject);
  private
    { Private declarations }
    //procedure Setdata;
    function MakeTempFileName: string;
    function GetHYID(sKH:string):integer;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.BtnSJDRClick(Sender: TObject);
var
  ExcelApp: Variant;
  Loop: integer;
  filename, sName: string;
  tmpstr: string;
begin
  if edtTBL.Text='' then
  begin
    showmessage('请输入表名');
    exit;
  end;
  {if edtZD.Text='' then
  begin
    showmessage('请输入导入字段');
    exit;
  end; }
  showmessage('所导入的excel文件数据自第二行起，第一列为卡号');
  OpenDialog1.FileName := '';
  OpenDialog1.Execute;
  if OpenDialog1.FileName = '' then
  begin
    //OpenDialog1.Free;
    exit;
  end;
  filename := OpenDialog1.FileName;

  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Visible := False;

  ExcelApp.Caption := '应用程序调用 Microsoft Excel';
  ExcelApp.WorkBooks.Add;
  ExcelApp.WorkBooks.open(filename);

  //Setdata;
  tblExcel.Close;
  tblExcel.EmptyTable;
  tblExcel.Open;

  for loop := 1 to ExcelApp.rows.count do
  begin
    tmpstr := ExcelApp.Cells[1 + loop, 1].value;
    if tmpstr = '' then
      break
    else
    begin
      with qryExcel do
      begin
        close;
        sql.Clear;
        sql.Text:='select 1 from HYK_HYXX where HYK_NO=:HYK_NO';
        ParamByName('HYK_NO').AsString:=tmpstr;
        open;
      end;
      if qryExcel.RecordCount<=0 then
      begin
        showmessage('卡号'+tmpstr+'在会员卡信息中不存在！请检查导入文件');
        tblExcel.Close;
        tblExcel.EmptyTable;
        tblExcel.Open;
        exit;
      end
      else
      begin
        if not tblExcel.Locate('HYK_NO',tmpstr,[]) then
        begin
          tblExcel.Append;
          tblExcel.FieldByName('HYK_NO').AsString := ExcelApp.Cells[1 + loop, 1].Value;
          tblExcel.Post;
        end;
      end;
    end;
  end;
  ExcelApp.WorkBooks.Close;
  ExcelApp.Quit;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  ConnectDB(Database1,'DB.ini');
  tblExcel.TableName:='xx';
  tblExcel.CreateTable;
  tblExcel.Open;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  if tblExcel.Exists then
  begin
    tblExcel.close;
    tblExcel.DeleteTable;
  end;
end;

function TForm1.GetHYID(sKH: string): integer;
begin
  with qryID do
  begin
    close;
    sql.Clear;
    sql.Text:='select HYID from HYK_HYXX where HYK_NO=:HYK_NO';
    ParamByName('HYK_NO').AsString:=sKH;
    open;
    if not IsEmpty then
      result:=Fields[0].AsInteger
    else
      result:=-1;
    close;
  end;
end;

function TForm1.MakeTempFileName: string;
var
  Hour, Min, Sec, MSec: Word;
begin
  DecodeTime(Now, Hour, Min, Sec, MSec);
  Result := Format('_%.2d%.2d%.2d%.4d', [Hour, Min, Sec, MSec]);
end;

{procedure TForm1.Setdata;
var
  Field: TField;
  i:integer;
  j:integer;
  sZD,sTmp:string;
begin
  if tblExcel.Exists then
  begin
    tblExcel.close;
    tblExcel.Destroy;
  end;
  tblExcel.TableName:=MakeTempFileName;
  //tblExcel.CreateTable;
  //tblExcel.Open;

  if edtZD.Text<>'' then
  begin
    sTmp:=edtZD.Text;
    for i:=0 to maxint do
    begin
      j:=pos(',',sTmp);
      if j>0 then
      begin
        sZD:=copy(sTmp,1,j-1);
        Field := TStringField.Create(nil);
        Field.Size := 20;
        field.FieldName := sZD;
        Field.DataSet := tblExcel;
        Field.DisplayLabel := sZD;
      end
      else
      begin
        sZD:=sTmp;
        Field := TStringField.Create(nil);
        Field.Size := 20;
        field.FieldName := sZD;
        Field.DataSet := tblExcel;
        Field.DisplayLabel := sZD;
        break;
      end;
      sTmp:=copy(sTmp,j+1,maxint);
      if trim(sTmp)='' then
        break;
    end;
    tblExcel.CreateTable;
    tblExcel.Open;
  end;
end;}

procedure TForm1.RzButton1Click(Sender: TObject);
begin
  qryConnect.Close;
  qryConnect.Open;
  Database1.StartTransaction();
  try
    with qryExcel do
    begin
      close;
      sql.Clear;
      sql.Text:='delete from HYK_LPFFTMPJL';
      execsql;
    end;

    tblExcel.First;
    while not tblExcel.Eof do
    begin
      with qryExcel do
      begin
        close;
        sql.Clear;
        sql.Text:='insert into HYK_LPFFTMPJL(HYID,HYK_NO)';
        sql.add('values(:HYID,:HYK_NO)');
        ParamByName('HYK_NO').AsString:=tblExcelHYK_NO.AsString;
        ParamByName('HYID').AsInteger:=GetHYID(tblExcelHYK_NO.AsString);
        execsql;
      end;
      tblExcel.Next;
    end;
    Database1.Commit;
    qryConnect.Close;
    showmessage('写入完成！');
  except
    Database1.Rollback;
    ApplicationHandleException(Self);
    qryConnect.Close;
    exit;
  end;
end;

end.
