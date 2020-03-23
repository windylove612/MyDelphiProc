unit fuQImport3License;
{$I QImport3VerCtrl.Inc}
interface

uses
  {$IFDEF VCL16}
    Winapi.Windows,
    Winapi.Messages,
    System.SysUtils,
    System.Classes,
    Vcl.Graphics,
    Vcl.Controls,
    Vcl.Forms,
    Vcl.Dialogs,
    Vcl.StdCtrls,
    Vcl.ComCtrls;
  {$ELSE}
    Windows,
    Messages,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    StdCtrls,
    ComCtrls;
  {$ENDIF}

type
  TfmQImport3License = class(TForm)
    reLicense: TRichEdit;
    btnPrint: TButton;
    btnOK: TButton;
    procedure btnPrintClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmQImport3License: TfmQImport3License;

implementation

{$R *.dfm}

procedure TfmQImport3License.btnPrintClick(Sender: TObject);
begin
  reLicense.Print(Caption);
end;

procedure TfmQImport3License.FormCreate(Sender: TObject);
var
  resStream: TResourceStream;
begin
  if FindResource(HInstance, 'qi_eula', RT_RCDATA) > 0 then
  begin
    resStream := TResourceStream.Create(HInstance, 'qi_eula', RT_RCDATA);
    try
      reLicense.Lines.LoadFromStream(resStream);
    finally
      resStream.Free;
    end;
  end;
end;

end.
