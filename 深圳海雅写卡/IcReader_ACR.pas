unit IcReader_ACR;
                                 
interface
    Function ACR110_Open(ReaderPort:Integer):smallint; stdcall;external 'ACR110U.dll';  //打开串口 0成功 ReaderPort 从1到8
    Function ACR110_Close(hReader:smallint):smallint; stdcall;external 'ACR110U.dll';  //关闭
    Function ACR110_Reset(hReader:longint):longint; stdcall;external 'ACR110U.dll';  //复位
    //类型  0x02，长度 4,7,10，序列号     (m1卡的长度为4)
    Function ACR110_Select(hReader:smallint;pResultTagType:PByte;pResultTagLength:PByte;pResultSN:PByte):smallint; stdcall;external 'ACR110U.dll';  //选卡
    //扇区号，密码验证类型 0xAA,0xBB  0xFD, 0xAA是0，0xAF是0xAF，0xBF是0xBF ,ab的密码长度是6，其它的为空
    Function ACR110_Login(hReader:smallint;Sector:Byte;KeyType:Byte;StoredNo:Byte;pKey:PByte):longint; stdcall;external 'ACR110U.dll';  //登陆
    //块号，块数据     块号是偏移地址
    Function ACR110_Read(hReader:smallint;Block:Byte;pBlockData:PByte):smallint; stdcall;external 'ACR110U.dll';  //读卡
    //块号，块数据
    Function ACR110_Write(hReader:smallint;Block:Byte;pBlockData:PByte):smallint; stdcall;external 'ACR110U.dll';  //写卡
implementation
    {Function ACR110_Open; far external 'ACR110U.dll';
    Function ACR110_Close; far external 'ACR110U.dll';
    Function ACR110_Reset; far external 'ACR110U.dll';
    Function ACR110_Select; far external 'ACR110U.dll';
    Function ACR110_Login; far external 'ACR110U.dll';
    Function ACR110_Read; far external 'ACR110U.dll';
    Function ACR110_Write; far external 'ACR110U.dll'; }
end.
