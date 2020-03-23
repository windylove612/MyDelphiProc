function GetDecrypt_DES(pCODE:PChar;var pREAL,pMsg:PChar):boolean;stdcall;
名称：解密暗码
传入：加密的暗码（100）
返回：实际的明码（100），提示信息（100）