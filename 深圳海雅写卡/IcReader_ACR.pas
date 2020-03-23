unit IcReader_ACR;
                                 
interface
    Function ACR110_Open(ReaderPort:Integer):smallint; stdcall;external 'ACR110U.dll';  //�򿪴��� 0�ɹ� ReaderPort ��1��8
    Function ACR110_Close(hReader:smallint):smallint; stdcall;external 'ACR110U.dll';  //�ر�
    Function ACR110_Reset(hReader:longint):longint; stdcall;external 'ACR110U.dll';  //��λ
    //����  0x02������ 4,7,10�����к�     (m1���ĳ���Ϊ4)
    Function ACR110_Select(hReader:smallint;pResultTagType:PByte;pResultTagLength:PByte;pResultSN:PByte):smallint; stdcall;external 'ACR110U.dll';  //ѡ��
    //�����ţ�������֤���� 0xAA,0xBB  0xFD, 0xAA��0��0xAF��0xAF��0xBF��0xBF ,ab�����볤����6��������Ϊ��
    Function ACR110_Login(hReader:smallint;Sector:Byte;KeyType:Byte;StoredNo:Byte;pKey:PByte):longint; stdcall;external 'ACR110U.dll';  //��½
    //��ţ�������     �����ƫ�Ƶ�ַ
    Function ACR110_Read(hReader:smallint;Block:Byte;pBlockData:PByte):smallint; stdcall;external 'ACR110U.dll';  //����
    //��ţ�������
    Function ACR110_Write(hReader:smallint;Block:Byte;pBlockData:PByte):smallint; stdcall;external 'ACR110U.dll';  //д��
implementation
    {Function ACR110_Open; far external 'ACR110U.dll';
    Function ACR110_Close; far external 'ACR110U.dll';
    Function ACR110_Reset; far external 'ACR110U.dll';
    Function ACR110_Select; far external 'ACR110U.dll';
    Function ACR110_Login; far external 'ACR110U.dll';
    Function ACR110_Read; far external 'ACR110U.dll';
    Function ACR110_Write; far external 'ACR110U.dll'; }
end.
