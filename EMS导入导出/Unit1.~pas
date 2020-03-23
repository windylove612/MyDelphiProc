unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, DBTables, RzButton, QImport3, QImport3Xlsx,fuQImport3XlsxEditor;

type
  TForm1 = class(TForm)
    RzButton1: TRzButton;
    Table1: TTable;
    Table1HYK_NO: TStringField;
    QImport3Xlsx1: TQImport3Xlsx;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure RzButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  Table1.TableName:='aa';
  Table1.CreateTable;
  Table1.Open;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  if Table1.Exists then
  begin
    Table1.DeleteTable;
  end;
end;

procedure TForm1.RzButton1Click(Sender: TObject);
begin
  //if TQImport2XLS(QImport2XLS1) then
  //QImport2XLS1.Execute;
  if RunQImportXlsxEditor(QImport3Xlsx1) then
  begin
    QImport3Xlsx1.Execute;
  end;
end;

end.
