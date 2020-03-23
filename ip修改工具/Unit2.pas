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


//���IP�Ƿ�Ϸ�
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
//��ȡCONFIG�е���Ϣ
  Form2.edt1.Text := inifile.ReadString('����','�������','');
  Form2.edt2.Text := inifile.ReadString('����','IP��','');
  Form2.edt3.Text := inifile.ReadString('����','��������','');
  Form2.edt4.Text := inifile.ReadString('����','Ĭ������','');
  Form2.edt5.Text := inifile.ReadString('����','DNS','');
  Form2.edt6.Text := inifile.ReadString('����','EndIP','');
  Form2.edt7.Text := inifile.ReadString('����','ƫ��','');
//��CHECKED��ʼ
  I := inifile.ReadString('����','DHCP','');
  I1 := inifile.ReadString('����','IPX','');
  I2 := inifile.ReadString('����','CSKEY','');
  I3 := inifile.ReadString('����','ipauto','');
  I4 := inifile.ReadString('����','IPƫ��','');
  
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

//����
end;

function GetCheck(str:string):Boolean;
begin

end;

procedure TForm2.btn1Click(Sender: TObject);
begin
//�����͹ر���ع��ܴ��뿪ʼ
  if Self.chk1.Checked = True   then
  inifile.WriteString('����','DHCP','Yes')
  else
  inifile.WriteString('����','DHCP','NO');
  if Self.chk2.Checked = True   then
  inifile.WriteString('����','IPX','Yes')
  else
  inifile.WriteString('����','IPX','NO');
  if Self.chk3.Checked = True then
  inifile.WriteString('����','CSKEY','Yes')
  else
  inifile.WriteString('����','CSKEY','NO');
  if Self.chk4.Checked = True then
  inifile.WriteString('����','ipauto','Yes')
  else
  inifile.WriteString('����','ipauto','NO');
  if Self.CheckBox1.Checked = False then
  inifile.WriteString('����','IPƫ��','NO')
  else
  inifile.WriteString('����','IPƫ��','Yes');
//�������
//����Ƿ������ַ����뿪ʼ
  if  Form2.edt1.Text =''  then
  begin
    Application.MessageBox('δ���ü������,��������Ҫ���õļ������!', '��ʾ',
      MB_OK + MB_ICONINFORMATION);
      Exit;
    end;
  if  Form2.edt2.Text = ''  then
  begin
        Application.MessageBox('��ʼIP�����ڲ���Ϊ��,��������Ҫ���õĿ�ʼIP��!', '��ʾ',
      MB_OK + MB_ICONINFORMATION);
      Exit;
    end;
  if Form2.edt6.Text = '' then
  begin
    Application.MessageBox('����IP�����ڲ���Ϊ��,���ֻ��һ��IP�Σ���������Ϳ�ʼIP����ͬ��IP', '��ʾ',
      MB_OK + MB_ICONINFORMATION);
      Exit;
    end;
  if  Form2.edt3.Text =''  then
  begin
        Application.MessageBox('�������������ݲ���Ϊ��,��������ȷ����������!', '��ʾ',
      MB_OK + MB_ICONINFORMATION);
      Exit;
    end;
  if  Form2.edt4.Text =''  then
  begin
        Application.MessageBox('���������ݲ���Ϊ��,��������ȷ������!', '��ʾ',
      MB_OK + MB_ICONINFORMATION);
      Exit;
    end;
  if  Form2.edt5.Text = ''  then
  begin
        Application.MessageBox('DNS�����ݲ���Ϊ��,��������ȷ��DNS!', '��ʾ',
      MB_OK + MB_ICONINFORMATION);
      Exit;
    end;
    if  CheckBox1.Checked  then
    begin
      if edt7.Text <= '0' then
      begin
        Application.MessageBox('���Ѿ�����IP��ַƫ�ƹ�����������Ҫƫ�Ƶ�ֵ', '��ʾ',
      MB_OK + MB_ICONINFORMATION);
      edt7.SetFocus;
      Exit;
      end;
    end;
//������
  if CheckIp(Form2.edt2.Text) then
  begin
    inifile.WriteString('����','IP��',Form2.edt2.Text);
    if CheckIp(Form2.edt6.Text) then
    begin
       inifile.WriteString('����','EndIP',Form2.edt6.Text);
       if Checkmask(Form2.edt4.Text) then
       begin
         inifile.WriteString('����','Ĭ������',Form2.edt4.Text);
         if Checkmask(Form2.edt3.Text) then
           begin
           inifile.WriteString('����','��������',Form2.edt3.Text);
           if Form2.edt1.Text <> '' then
             begin
             inifile.WriteString('����','�������',Form2.edt1.Text);
             if CheckBox1.Checked =False then
             begin
             inifile.WriteString('����','ƫ��','0')
             end else
              inifile.WriteString('����','ƫ��',Form2.edt7.Text);
              inifile.WriteString('����','DNS',Form2.edt5.Text);
             Application.MessageBox('�޸����ò����ɹ�!', '��ʾ', MB_OK +
               MB_ICONINFORMATION) ;
           end else
               Application.MessageBox('���ò����޸�ʧ��,���������ļ��Ƿ����!',
               '��ʾ', MB_OK + MB_ICONINFORMATION);
         end else
           Application.MessageBox('����������д����������������ȷ�������룡',
             '����', MB_OK + MB_ICONSTOP);
       end else
         Application.MessageBox('Ĭ��������д����������������ȷ���ص�ַ���룡',
           '����', MB_OK + MB_ICONSTOP);
    end else
      Application.MessageBox('����IP����д�������������룡', '����', MB_OK +
        MB_ICONSTOP);
  end else
    Application.MessageBox('��ʼIP����д�������������룡', '����', MB_OK +
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
//���IP�Ƿ�Ϸ��ķ�ʽ
//    tmp := TStringList.Create;
//    str := inifile.ReadString('����','IP��','');
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
//             ShowMessage('IP��ַ����')
//           end  else
//           ShowMessage('IP��ַ����')
//         end else
//         ShowMessage('IP��ַ����')
//       end else
//       ShowMessage('IP��ַ����')
//      end else
//      ShowMessage('IP��ַ����')
////    ShowMessage('��ȷ');
//    end else
//    ShowMessage('IP��ַ����');
//���IP�Ƿ�Ϸ��ķ�ʽ����

end.
