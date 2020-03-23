unit fmUltraPlatForm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, RzPanel, ImgList, RzCommon;

type
  TForm1 = class(TForm)
    RzFrameController: TRzFrameController;
    ButtonImage: TImageList;
    RzPanel1: TRzPanel;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

end.
