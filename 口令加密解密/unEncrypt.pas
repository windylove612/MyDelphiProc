unit unEncrypt;

interface

function Encrypt(const S: String): String;
function Decrypt(const S: String): String;

implementation

//const  C1 = 52845;  C2 = 22719;
//const  KeyWord = 7891;
const  C1 = 21469;  C2 = 12347;
const  KeyWord = 26493;

{$R-}
function Encrypt(const S: String): String;
var
  I: byte;
  Key: Word;
begin
  Key := Keyword;
  SetLength(Result,Length(S));
  for I := 1 to Length(S) do
  begin
    Result[I] := char(byte(S[I]) xor (Key shr 8));
    Key := (byte(Result[I]) + Key) * C1 + C2;
  end;
end;

function Decrypt(const S: String): String;
var
  I: byte;
  Key: Word;
begin
  Key := Keyword;
  SetLength(Result,Length(S));
  for I := 1 to Length(S) do
  begin
    Result[I] := char(byte(S[I]) xor (Key shr 8));
    Key := Word((byte(S[I]) + Key) * C1 + C2);
  end;
end;

end.
