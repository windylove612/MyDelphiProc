unit fmProcWriteICK;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,IcReader_ACR, DB, DBTables, StdCtrls, RzEdit, RzButton,
  Grids, DBGrids, RzDBGrid, RzLabel, ExtCtrls, RzPanel,comobj,fmSetupPort_ACR,
  BMPBTN;

const
  byKey: array[0..5] of Byte = ($37, $61, $34, $54, $66, $47);
  key: array[0..5] of Char = (Char($37), Char($61), Char($34), Char($54), Char($66), Char($47));
  pKey: array[0..5] of Byte = ($37, $61, $34, $54, $66, $47);

type
  TProcZK = class(TForm)
    RzPanel1: TRzPanel;
    RzDBGrid1: TRzDBGrid;
    RzButton1: TRzButton;
    RzButton2: TRzButton;
    tblMain: TTable;
    dsMain: TDataSource;
    tblMainHYK_NO: TStringField;
    tblMainCDNR: TStringField;
    RzButton3: TRzButton;
    RzLabel2: TRzLabel;
    OpenDialog1: TOpenDialog;
    RzPanel2: TRzPanel;
    RzLabel1: TRzLabel;
    RzLabel3: TRzLabel;
    lblKH: TRzLabel;
    lblCD: TRzLabel;
    btnSetup: TBmpBtn;
    Label7: TLabel;
    lblPortSetting: TLabel;
    Bevel1: TBevel;
    Label10: TLabel;
    Label11: TLabel;
    LabXK: TLabel;
    LabSY: TLabel;
    Label12: TLabel;
    LabSBSL: TLabel;
    RzLabel4: TRzLabel;
    lblDCKH: TRzLabel;
    RzLabel6: TRzLabel;
    lblDCCD: TRzLabel;
    procedure RzButton3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure dsMainDataChange(Sender: TObject; Field: TField);
    procedure btnSetupClick(Sender: TObject);
    procedure RzButton1Click(Sender: TObject);
    procedure RzButton2Click(Sender: TObject);
  private
    { Private declarations }
    pSetting: TUSBPortSetting;
    ICDEV     : Longint;
    Comm : string;
    iXK : Integer;
    iSBSL : Integer;
    st:integer;
    Function ReadTextFromTextFile(sTextFile:String):String;
    Procedure GetFromTextFile(sLineText:String);
    function MakeTempFileName: string;
    procedure ShowPortSetting;
    function getCardID(var tmpCardID:string):boolean;
    function IniPassWord(s,sCZKHM,sCD:string):boolean;
    function  btnCkClick:boolean;
  public
    { Public declarations }
  end;

var
  ProcZK: TProcZK;

implementation

{$R *.dfm}

{ TProcZK }

procedure TProcZK.GetFromTextFile(sLineText: String);
var
  iPos1,iPos2,i:integer;
  TmpStr,Str : String;
  sCD:string;
begin
  Str:=sLineText;
  tblMain.Append;
  iPos1 := Pos(',',Str);
  try
    TmpStr := Copy(Str,1,iPos1-1);
    sCD := Copy(Str,iPos1+1,Length(Str)-iPos1);
    tblMainHYK_NO.AsString := Trim(TmpStr);
    tblMainCDNR.AsString := Trim(sCD);
  except
    On E:Exception do
    begin
      ShowMessage('导入的文本文件格式有错：'+E.Message);
    end;
  end;
  tblMain.Post;
end;

function TProcZK.MakeTempFileName: string;
var
  Hour, Min, Sec, MSec: Word;
begin
  DecodeTime(Now, Hour, Min, Sec, MSec);
  Result := Format('_%.2d%.2d%.2d%.4d', [Hour, Min, Sec, MSec]);
end;

function TProcZK.ReadTextFromTextFile(sTextFile: String): String;
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

procedure TProcZK.RzButton3Click(Sender: TObject);
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
  tblMain.Close;
  tblMain.EmptyTable;
  tblMain.Open;
  ReadTextFromTextFile(filename);
  tblMain.First;
  iSBSL := 0;
  iXK:=0;
  LabXK.Caption:=IntToStr(iXK);;
  LabSBSL.Caption := IntToStr(iSBSL);
  LabSY.Caption := IntToStr(tblMain.RecordCount);
end;

procedure TProcZK.FormCreate(Sender: TObject);
begin
  tblMain.TableName:=MakeTempFileName;
  tblMain.CreateTable;
  tblMain.Open;
end;

procedure TProcZK.FormDestroy(Sender: TObject);
begin
  if tblMain.Exists then
  begin
    tblMain.close;
    tblMain.DeleteTable;
  end;
  ACR110_Close(pSetting.iPort);
end;

procedure TProcZK.dsMainDataChange(Sender: TObject; Field: TField);
begin
  lblKH.Caption:=tblMainHYK_NO.AsString;
  lblCD.Caption:=tblMainCDNR.AsString;
end;

procedure TProcZK.btnSetupClick(Sender: TObject);
begin
  if DoSetPort_ACR(Self, pSetting) then
    ShowPortSetting();

  icdev := ACR110_Open(pSetting.iPort);//pSetting.iPort
  if icdev < 0 then      //初始化设备 ...................... . (1)
  begin
    ShowMessage('读写器连接失败,请检查设备！');
    exit;
  end;
end;

procedure TProcZK.ShowPortSetting;
var
  tmpstr1	:string;
begin
  tmpstr1 := Format('端口设置: %s',['USB端口:'+IntToStr(pSetting.iPort)]);
  lblPortSetting.Caption := tmpstr1;
  Comm := 'USB端口:'+IntToStr(pSetting.iPort);
end;

procedure TProcZK.RzButton1Click(Sender: TObject);
var
  sICKHM,yICKHM,sPassWord,tmpStr:String;
  sICK16HM:String;
  hexkey: pchar;
  rebuff: array[0..31] of char;
  sData:pChar;
  i: Integer;
  sKH:string;
  tmpCDNR:String;
BEgin
  sICKHM := '';
  yICKHM := '';
  sICK16HM :='';
  sPassWord := '';

  if not getCardID(sICKHM) then     //核对密码取卡序列号给sICKHM
  begin
    if not getCardID(sICKHM) then
    begin
      if not getCardID(sICKHM) then
      begin
        ShowMessage('写卡--读IC卡序列号失败！');
        Exit;
      end;
    end;
  end;

  TmpStr := Trim(lblKH.Caption);
  tmpCDNR := Trim(lblCD.Caption);
  try
    sPassWord:=sICKHM;    //卡序列号
    if not IniPassWord(sPassWord,tmpStr,tmpCDNR) then     //写卡密码
      raise Exception.Create('写卡--写卡失败');

    if not btnCkClick then
      raise Exception.Create('读卡--读卡失败，重新写卡！');

    Inc(iXK);
    tblMain.Delete();
    lblKH.Caption := tblMainHYK_NO.AsString;
    LabXK.Caption := IntToStr(iXK);
    LabSY.Caption := IntToStr(tblMain.RecordCount);
    LabSBSL.Caption := IntToStr(iSBSL);
    if tblMain.RecordCount = 0 then
    begin
      ShowMessage('本批卡写卡完毕！');
    end;
  except
    on E:Exception do
    begin
      ShowMessage('制卡失败，请与系统管理员联系，错误信息：'+E.Message);
      Exit;
    end;
  end;
end;

function TProcZK.getCardID(var tmpCardID: string): boolean;
var
  tempint: longword;
  address: Integer;
  sector,i: Integer;
  sTagType: array[0..49] of Byte;
  sTagLength: array[0..9] of Byte;
  sXLH: array[0..15] of Byte;
begin
  Result:= False;
  ZeroMemory(@sTagType, 50);
  ZeroMemory(@sTagLength, 10);
  ZeroMemory(@sXLH, 16);
  sector := 0;
  tmpCardID := '';
  if ACR110_Select(pSetting.iPort,@sTagType,@sTagLength,@sXLH)=0 then  //
  begin
    for i:=0 to 15 do
    begin
      tmpCardID := tmpCardID + char(sXLH[i]);
    end;
    tmpCardID := trim(tmpCardID);
    Result := True;
  end;
end;

function TProcZK.btnCkClick: boolean;
var
  iCount :integer;
  tmpdata :string;
  sCDNR:string;
  s16CDNR:string;
  rebuff: array[0..31] of char;
  sCD1,sCD2:array[0..15] of Byte;
  sKH:array[0..15] of Byte;
  sCD:string;
  s:string;
  i:integer;
begin
  s16CDNR:='';
  Result :=false;
  if not getCardID(sCDNR) then
  begin
    if not getCardID(sCDNR) then
    begin
      if not getCardID(sCDNR) then
      begin
        ShowMessage('读卡--读IC卡失败！');
        Exit;
      end;
    end;
  end;

  ZeroMemory(@sCD1, 16);
  ZeroMemory(@sCD2, 16);
  ZeroMemory(@sKH, 16);

  if ACR110_Login(pSetting.iPort,$05, $AA, $02,@pKey)<>0 then
  begin
    showmessage('登陆5扇区失败');
    exit;
  end;

  st:=ACR110_Read(pSetting.iPort,22,@sKH);          //读取磁道，二扇区的1，2块存磁道，用f来分隔
  if st<>0 then
  begin
     showMessage('已写入卡读卡-读取数据失败!') ;
     exit;
  end;

  for i := 0 to 15 do
  begin
    sCD := sCD + char(sKH[i]);
  end;

  lblDCKH.Caption:=trim(sCD);

  if ACR110_Login(pSetting.iPort,$06, $AA, $00,@pKey)<>0 then
  begin
    showmessage('登陆2扇区失败');
    exit;
  end;

  st:=ACR110_Read(pSetting.iPort,24,@sCD1);          //读取磁道，二扇区的1，2块存磁道，用f来分隔
  if st<>0 then
  begin
     showMessage('已写入卡读卡-读取数据失败!') ;
     exit;
  end;

  st:=ACR110_Read(pSetting.iPort,25,@sCD2);          //读取磁道，二扇区的1，2块存磁道，用f来分隔
  if st<>0 then
  begin
     showMessage('已写入卡读卡-读取数据失败!') ;
     exit;
  end;

  for i := 0 to 15 do
  begin
    sCD := sCD + char(sCD1[i]);
  end;

  for i := 0 to 15 do
  begin
    sCD := sCD + char(sCD2[i]);
  end;

  lblDCCD.Caption:=trim(sCD);

  result :=true;  
end;

function TProcZK.IniPassWord(s, sCZKHM, sCD: string): boolean;
var
  tmpPswd,TmpStr1:string;
  sXLH,i : Integer ;
  XLHint: longword;
  sCD1:array[0..15] of Byte;
  sCD2:array[0..15] of Byte;
  sKH:array[0..15] of Byte;
  iCDNR:pchar;
  sKHCD:integer;
begin
  Result := False;
  tmpPswd := s;

  //写扇区密码
  if ACR110_WriteMasterKey(pSetting.iPort,$05,@pKey)<>0 then
  begin
    showmessage('写5扇区密码失败');
    exit;
  end;

  if ACR110_Login(pSetting.iPort,$05, $AA, $02,@pKey)<>0 then
  begin
    showmessage('登陆5扇区失败');
    exit;
  end;

  sKHCD:=length(trim(sCZKHM));

  for i := 0 to 15 do
  begin
    sKH[i]:=$20;
  end;

  if ACR110_Write(pSetting.iPort,22,@sKH)<>0 then   //写卡号到5扇区2块
  begin
    if ACR110_Write(pSetting.iPort,22,@sKH)<>0 then
    begin
      if ACR110_Write(pSetting.iPort,22,@sKH)<>0 then
      begin
        ShowMessage('写入5扇区信息失败！');
        Exit ;
      end;
    end;
  end;

  //写扇区密码
  if ACR110_WriteMasterKey(pSetting.iPort,$06,@pKey)<>0 then
  begin
    showmessage('写6扇区密码失败');
    exit;
  end;

  if ACR110_Login(pSetting.iPort,$06, $AA, $00,@pKey)<>0 then
  begin
    showmessage('登陆6扇区失败');
    exit;
  end;

  sKHCD:=length(trim(sCD));

  if sKHCD<16 then
  begin
    for i:=0 to sKHCD-1 do
    begin
      if copy(trim(sCD),i+1,1)<>'=' then
        sCD1[i]:=StrToInt64(copy(trim(sCD),i+1,1))+48
      else
        sCD1[i]:=61;
    end;
    for i := sKHCD to 15 do
    begin
      sCD1[i]:=$20;
    end;
    for i := 0 to 15 do
    begin
      sCD2[i]:=$20;
    end;
  end
  else if sKHCD=16 then
  begin
    for i:=0 to sKHCD-1 do
    begin
      if copy(trim(sCD),i+1,1)<>'=' then
        sCD1[i]:=StrToInt64(copy(trim(sCD),i+1,1))+48
      else
        sCD1[i]:=61;
    end;
    for i := 0 to 15 do
    begin
      sCD2[i]:=$20;
    end;
  end
  else
  begin
    for i:=0 to 15 do
    begin
      if copy(trim(sCD),i+1,1)<>'=' then
        sCD1[i]:=StrToInt64(copy(trim(sCD),i+1,1))+48
      else
        sCD1[i]:=61;
    end;
    for i:=0 to sKHCD-17 do
    begin
      if copy(trim(sCD),i+17,1)<>'=' then
        sCD2[i]:=StrToInt64(copy(trim(sCD),i+17,1))+48
      else
        sCD2[i]:=61;
    end;
    if sKHCD<32 then
    begin
      for i := sKHCD-16 to 15 do
      begin
        sCD2[i]:=$20;
      end;
    end;
  end;

  if sKHCD<16 then
  begin
    if ACR110_Write(pSetting.iPort,24,@sCD1)<>0 then   //写磁道到二扇区1块、2块
    begin
      if ACR110_Write(pSetting.iPort,24,@sCD1)<>0 then
      begin
        if ACR110_Write(pSetting.iPort,24,@sCD1)<>0 then
        begin
          ShowMessage('写入磁道信息失败！');
          Exit ;
        end;
      end;
    end;
  end
  else if sKHCD=16 then
  begin
    if ACR110_Write(pSetting.iPort,24,@sCD1)<>0 then   //写磁道到二扇区1块、2块
    begin
      if ACR110_Write(pSetting.iPort,24,@sCD1)<>0 then
      begin
        if ACR110_Write(pSetting.iPort,24,@sCD1)<>0 then
        begin
          ShowMessage('写入磁道信息失败！');
          Exit ;
        end;
      end;
    end;
  end
  else
  begin
    if ACR110_Write(pSetting.iPort,24,@sCD1)<>0 then   //写磁道到二扇区1块、2块
    begin
      if ACR110_Write(pSetting.iPort,24,@sCD1)<>0 then
      begin
        if ACR110_Write(pSetting.iPort,24,@sCD1)<>0 then
        begin
          ShowMessage('写入磁道信息失败！');
          Exit ;
        end;
      end;
    end;
    if ACR110_Write(pSetting.iPort,25,@sCD2)<>0 then   //写磁道到二扇区1块、2块
    begin
      if ACR110_Write(pSetting.iPort,25,@sCD2)<>0 then
      begin
        if ACR110_Write(pSetting.iPort,25,@sCD2)<>0 then
        begin
          ShowMessage('写入磁道信息失败！');
          Exit ;
        end;
      end;
    end;
  end;

  sleep(5);

  Result := True;
end;

procedure TProcZK.RzButton2Click(Sender: TObject);
var
  iCount :integer;
  tmpdata :string;
  sCDNR:string;
  s16CDNR:string;
  rebuff: array[0..31] of char;
  sCD1,sCD2:array[0..15] of Byte;
  sKH:array[0..15] of Byte;
  sCD:string;
  i:integer;
  HaveTag: array[0..49] of Byte;
  tmpArray: array[0..9] of Byte;
  dataRead: array[0..15] of Byte;
begin
  sCD:='';
  if not getCardID(sCDNR) then
  begin
    if not getCardID(sCDNR) then
    begin
      if not getCardID(sCDNR) then
      begin
        ShowMessage('读卡--读IC卡失败！');
        Exit;
      end;
    end;
  end;

  ZeroMemory(@sCD1, 16);
  ZeroMemory(@sCD2, 16);
  ZeroMemory(@sKH, 16);

  if ACR110_Login(pSetting.iPort,$05,$AA,$02,@pKey)<>0 then
  begin
    showmessage('登陆5扇区失败');
    exit;
  end;

  st:=ACR110_Read(pSetting.iPort,22,@sKH);          //读取磁道，二扇区的1，2块存磁道，用f来分隔
  if st<>0 then
  begin
     showMessage('已写入卡读卡-读取数据失败!') ;
     exit;
  end;

  for i := 0 to 15 do
  begin
    sCD := sCD + char(sKH[i]);
  end;

  lblDCKH.Caption:=trim(sCD);

  if ACR110_Login(pSetting.iPort,$06,$AA,$00,@pKey)<>0 then
  begin
    showmessage('登陆2扇区失败');
    exit;
  end;

  st:=ACR110_Read(pSetting.iPort,24,@sCD1);          //读取磁道，二扇区的1，2块存磁道，用f来分隔
  if st<>0 then
  begin
     showMessage('已写入卡读卡-读取数据失败!') ;
     exit;
  end;

  st:=ACR110_Read(pSetting.iPort,25,@sCD2);          //读取磁道，二扇区的1，2块存磁道，用f来分隔
  if st<>0 then
  begin
     showMessage('已写入卡读卡-读取数据失败!') ;
     exit;
  end;

  for i := 0 to 15 do
  begin
    sCD := sCD + char(sCD1[i]);
  end;

  for i := 0 to 15 do
  begin
    sCD := sCD + char(sCD2[i]);
  end;

  lblDCCD.Caption:=trim(sCD);
end;

end.
