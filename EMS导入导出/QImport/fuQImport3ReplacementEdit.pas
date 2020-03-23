unit fuQImport3ReplacementEdit;
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
    Vcl.ExtCtrls;
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
    ExtCtrls;
  {$ENDIF}

type
  TfmQImport3ReplacementEdit = class(TForm)
    bOk: TButton;
    bCancel: TButton;
    laTextToFind: TLabel;
    laReplaceWith: TLabel;
    chIgnoreCase: TCheckBox;
    edTextToFind: TEdit;
    edReplaceWith: TEdit;
    Bevel: TBevel;
    procedure bOkClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

function ReplacementEdit(var TextToFind, ReplaceWith: string;
  var IgnoreCase: boolean): boolean;

implementation

{$R *.DFM}

uses QImport3StrIDs, QImport3;

function ReplacementEdit(var TextToFind, ReplaceWith: string;
  var IgnoreCase: boolean): boolean;
begin
  with TfmQImport3ReplacementEdit.Create(nil) do
    try
      edTextToFind.Text := TextToFind;
      edReplaceWith.Text := ReplaceWith;
      chIgnoreCase.Checked := IgnoreCase;

      Result := ShowModal = mrOk;

      if Result then begin
        TextToFind := edTextToFind.Text;
        ReplaceWith := edReplaceWith.Text;
        IgnoreCase := chIgnoreCase.Checked;
      end;
    finally
      Free;
    end;
end;

procedure TfmQImport3ReplacementEdit.bOkClick(Sender: TObject);
begin
  if edTextToFind.Text = EmptyStr then
    raise Exception.Create(QImportLoadStr(QIE_NeedDefineTextToFind));
  ModalResult := mrOk;
end;

procedure TfmQImport3ReplacementEdit.FormCreate(Sender: TObject);
begin
  Caption := QImportLoadStr(QIWDF_RE_Title);
  laTextToFind.Caption := QImportLoadStr(QIWDF_RE_TextToFind);
  laReplaceWith.Caption := QImportLoadStr(QIWDF_RE_ReplaceWith);
  chIgnoreCase.Caption := QImportLoadStr(QIWDF_RE_IgnoreCase);
end;

end.
