unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, inifiles;

type
  TForm2 = class(TForm)
    lbl1: TLabel;
    lbl2: TLabel;
    lbl3: TLabel;
    lbl4: TLabel;
    edt1: TEdit;
    edt2: TEdit;
    edt3: TEdit;
    edt4: TEdit;
    lbl5: TLabel;
    edt5: TEdit;
    chk1: TCheckBox;
    btn1: TButton;
    btn2: TButton;
    chk2: TCheckBox;
    chk3: TCheckBox;
    chk4: TCheckBox;
    lbl6: TLabel;
    edt6: TEdit;
    lbl7: TLabel;
    CheckBox1: TCheckBox;
    edt7: TEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btn2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btn1Click(Sender: TObject);


  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;
  form1 : TForm;
  inifile: tinifile;

implementation

uses Unit1;

{$R *.dfm}


procedure TForm2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//Application.Terminate;
end;

procedure TForm2.btn2Click(Sender: TObject);
begin
  Close;
end;


procedure TForm2.FormCreate(Sender: TObject);
begin
  inifile := tinifile.Create(ExtractFileDir(Application.exename) + '\config.ini');
end;


//检查IP是否合法
function CheckIp(sIP: string): Boolean;
var
  i,j: Integer;
  List: TStringList;
begin
  Result := False;
  List := TStringList.Create;
  List.Delimiter := '.';
  List.DelimitedText := sIP;
  if List.Count <> 4 then
  begin
    List.Free;
    Exit;
  end;
  for i := 0 to List.Count - 2 do
  begin
    j := StrToIntDef(List.Strings[i],0);
    if (j < 0) or (j > 254) then
    begin
      List.Free;
      Exit;
    end;
  end;
  List.Free;
  Result := True;
end;

function  Checkmask(sIP:string):Boolean;
var
 i,j: Integer;
  List: TStringList;
begin
  Result := False;
  List := TStringList.Create;
  List.Delimiter := '.';
  List.DelimitedText := sIP;
  if List.Count <> 4 then
  begin
    List.Free;
    Exit;
  end;
  for i := 0 to List.Count - 1 do
  begin
    j := StrToIntDef(List.Strings[i],0);
    if (j < 0) or (j > 255) then
    begin
      List.Free;
      Exit;
    end;
  end;
  List.Free;
  Result := True;
end;

procedure TForm2.FormShow(Sender: TObject);
var
  I,I1,I2,I3,I4 :string;

begin
//读取CONFIG中的信息
  Form2.edt1.Text := inifile.ReadString('配置','计算机名','');
  Form2.edt2.Text := inifile.ReadString('配置','IP段','');
  Form2.edt3.Text := inifile.ReadString('配置','子网掩码','');
  Form2.edt4.Text := inifile.ReadString('配置','默认网关','');
  Form2.edt5.Text := inifile.ReadString('配置','DNS','');
  Form2.edt6.Text := inifile.ReadString('配置','EndIP','');
  Form2.edt7.Text := inifile.ReadString('配置','偏移','');
//读CHECKED开始
  I := inifile.ReadString('配置','DHCP','');
  I1 := inifile.ReadString('配置','IPX','');
  I2 := inifile.ReadString('配置','CSKEY','');
  I3 := inifile.ReadString('配置','ipauto','');
  I4 := inifile.ReadString('配置','IP偏移','');
  
  if  I = 'Yes'   then
  Self.chk1.Checked := True
  else
  Self.chk1.Checked := False;
  if I1 ='Yes'   then
  Self.chk2.Checked := True
  else
  Self.chk2.Checked := False;
  if I2 = 'Yes' then
  Self.chk3.Checked := True
  else
  Self.chk3.Checked := False;
  if  I3 = 'Yes' then
  Self.chk4.Checked := True
  else
  Self.chk4.Checked := False ;
  if I4= 'Yes'  then
  Self.CheckBox1.Checked :=True
  else
  Self.CheckBox1.Checked := False;

//结束
end;

function GetCheck(str:string):Boolean;
begin

end;

procedure TForm2.btn1Click(Sender: TObject);
begin
//开启和关闭相关功能代码开始
  if Self.chk1.Checked = True   then
  inifile.WriteString('配置','DHCP','Yes')
  else
  inifile.WriteString('配置','DHCP','NO');
  if Self.chk2.Checked = True   then
  inifile.WriteString('配置','IPX','Yes')
  else
  inifile.WriteString('配置','IPX','NO');
  if Self.chk3.Checked = True then
  inifile.WriteString('配置','CSKEY','Yes')
  else
  inifile.WriteString('配置','CSKEY','NO');
  if Self.chk4.Checked = True then
  inifile.WriteString('配置','ipauto','Yes')
  else
  inifile.WriteString('配置','ipauto','NO');
  if Self.CheckBox1.Checked = False then
  inifile.WriteString('配置','IP偏移','NO')
  else
  inifile.WriteString('配置','IP偏移','Yes');
//代码结束
//检测是否输入字符代码开始
  if  Form2.edt1.Text =''  then
  begin
    Application.MessageBox('未设置计算机名,请输入你要设置的计算机名!', '提示',
      MB_OK + MB_ICONINFORMATION);
      Exit;
    end;
  if  Form2.edt2.Text = ''  then
  begin
        Application.MessageBox('开始IP段项内不能为空,请输入你要设置的开始IP段!', '提示',
      MB_OK + MB_ICONINFORMATION);
      Exit;
    end;
  if Form2.edt6.Text = '' then
  begin
    Application.MessageBox('结束IP段项内不能为空,如果只有一个IP段，请你输入和开始IP段相同的IP', '提示',
      MB_OK + MB_ICONINFORMATION);
      Exit;
    end;
  if  Form2.edt3.Text =''  then
  begin
        Application.MessageBox('子网掩码项内容不能为空,请输入正确的子网掩码!', '提示',
      MB_OK + MB_ICONINFORMATION);
      Exit;
    end;
  if  Form2.edt4.Text =''  then
  begin
        Application.MessageBox('网关项内容不能为空,请输入正确的网关!', '提示',
      MB_OK + MB_ICONINFORMATION);
      Exit;
    end;
  if  Form2.edt5.Text = ''  then
  begin
        Application.MessageBox('DNS项内容不能为空,请输入正确的DNS!', '提示',
      MB_OK + MB_ICONINFORMATION);
      Exit;
    end;
    if  CheckBox1.Checked  then
    begin
      if edt7.Text <= '0' then
      begin
        Application.MessageBox('你已经开启IP地址偏移功能请输入需要偏移的值', '提示',
      MB_OK + MB_ICONINFORMATION);
      edt7.SetFocus;
      Exit;
      end;
    end;
//检测结束
  if CheckIp(Form2.edt2.Text) then
  begin
    inifile.WriteString('配置','IP段',Form2.edt2.Text);
    if CheckIp(Form2.edt6.Text) then
    begin
       inifile.WriteString('配置','EndIP',Form2.edt6.Text);
       if Checkmask(Form2.edt4.Text) then
       begin
         inifile.WriteString('配置','默认网关',Form2.edt4.Text);
         if Checkmask(Form2.edt3.Text) then
           begin
           inifile.WriteString('配置','子网掩码',Form2.edt3.Text);
           if Form2.edt1.Text <> '' then
             begin
             inifile.WriteString('配置','计算机名',Form2.edt1.Text);
             if CheckBox1.Checked =False then
             begin
             inifile.WriteString('配置','偏移','0')
             end else
              inifile.WriteString('配置','偏移',Form2.edt7.Text);
              inifile.WriteString('配置','DNS',Form2.edt5.Text);
             Application.MessageBox('修改配置参数成功!', '提示', MB_OK +
               MB_ICONINFORMATION) ;
           end else
               Application.MessageBox('配置参数修改失败,请检查配置文件是否存在!',
               '提示', MB_OK + MB_ICONINFORMATION);
         end else
           Application.MessageBox('子网掩码填写错误，请重新输入正确子网掩码！',
             '错误', MB_OK + MB_ICONSTOP);
       end else
         Application.MessageBox('默认网关填写错误，请重新输入正确网关地址掩码！',
           '错误', MB_OK + MB_ICONSTOP);
    end else
      Application.MessageBox('结束IP段填写错误，请重新输入！', '错误', MB_OK +
        MB_ICONSTOP);
  end else
    Application.MessageBox('开始IP段填写错误，请重新输入！', '错误', MB_OK +
      MB_ICONSTOP);
end;

//procedure TForm2.btn3Click(Sender: TObject);
//var
//  str: string;
//  tmp: TStringList;
//  tmp1: string;
//  str1,str2,str3 :Integer;
//begin
//  if CheckIp(edt7.Text) then
//    ShowMessage('ok')
//  else
//    ShowMessage('error');
//检查IP是否合法的方式
//    tmp := TStringList.Create;
//    str := inifile.ReadString('配置','IP段','');
//    tmp.Text := StringReplace(str, '.', #13, [rfReplaceAll]);
////    ShowMessage(tmp.Strings[0]);
//    tmp1 := tmp.Strings[0] + '.' + tmp.Strings[1] + '.';
//    str1:=StrToInt(tmp.Strings[0]) ;
//    str2:= StrToInt(tmp.Strings[1]);
//    str3:= StrToInt(tmp.Strings[2]);
//  if str1 < 255  then
//  begin
//    if str1 > 0 then
//    begin
//     if  str2 < 255 then
//     begin
//       if str2> 0 then
//       begin
//         if str3 < 255 then
//         begin
//           if str3> 0 then
//           begin
//             ShowMessage('OK');
//             end else
//             ShowMessage('IP地址错误')
//           end  else
//           ShowMessage('IP地址错误')
//         end else
//         ShowMessage('IP地址错误')
//       end else
//       ShowMessage('IP地址错误')
//      end else
//      ShowMessage('IP地址错误')
////    ShowMessage('正确');
//    end else
//    ShowMessage('IP地址错误');
//检查IP是否合法的方式结束

end.
