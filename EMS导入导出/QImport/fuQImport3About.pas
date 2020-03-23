unit fuQImport3About;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    Vcl.Forms,
    Vcl.StdCtrls,
    Vcl.Buttons,
    Vcl.Controls,
    Vcl.ExtCtrls,
    System.Classes,
    Winapi.Windows,
    Vcl.Graphics;
  {$ELSE}
    Forms,
    StdCtrls,
    Buttons,
    Controls,
    ExtCtrls,
    Classes,
    Windows,
    Graphics;
  {$ENDIF}

type
  TfmQImport3About = class(TForm)
    Image: TImage;
    lbVerInfo: TLabel;
    lbCopyRight: TLabel;
    lbVersion: TLabel;
    BitBtn1: TButton;
    laWarn: TLabel;
    laDevelopers: TLabel;
    lbCompanyHomePageTag: TLabel;
    lbCompanyHomePage: TLabel;
    lbProductHomePageTag: TLabel;
    lbProductHomePage: TLabel;
    GroupBox1: TGroupBox;
    lbLicense: TLabel;
    Panel1: TPanel;
    lbRegisterNow: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure lbCompanyHomePageClick(Sender: TObject);
    procedure lbRegisterNowClick(Sender: TObject);
    procedure lbLicenseClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

procedure ShowAboutForm;

implementation

uses
  {$IFDEF VCL16}
    Winapi.ShellAPI,
    System.SysUtils,
  {$ELSE}
    ShellAPI,
    SysUtils,
  {$ENDIF}
  QImport3Common,
  fuQImport3License;

{$R *.DFM}

procedure ShowAboutForm;
begin
  with TfmQImport3About.Create(nil) do
  try
    lbVerInfo.Caption := Format(QI_FULL_PRODUCT_NAME, [QI_VERSION]);
    lbCopyRight.Caption := QI_COPYRIGHT;
    ShowModal;
  finally
    Free;
  end;
end;

procedure TfmQImport3About.FormCreate(Sender: TObject);
begin
{$IFDEF ADVANCED_DATA_IMPORT_TRIAL_VERSION}
  lbVersion.Caption := 'Evalution version';
{$ELSE}
  lbVersion.Caption := 'Registered version';
  lbVersion.Font.Color := clBlack;
  lbRegisterNow.Visible := False;
{$ENDIF}
end;

procedure TfmQImport3About.lbCompanyHomePageClick(Sender: TObject);
begin
  ShellExecute(Handle, 'open',
               PChar((Sender as TLabel).Caption),
               nil, nil, SW_SHOW);
end;

procedure TfmQImport3About.lbRegisterNowClick(Sender: TObject);
begin
  ShellExecute(Handle, 'open',
               PChar(QI_REG_URL),
               nil, nil, SW_SHOW);
end;

procedure TfmQImport3About.lbLicenseClick(Sender: TObject);
begin
  with TfmQImport3License.Create(nil) do
  try
    ShowModal;
  finally
    Free;
  end;
end;

end.
