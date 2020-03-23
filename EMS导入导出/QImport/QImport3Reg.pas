unit QImport3Reg;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    System.SysUtils,
    System.Classes,
  {$ELSE}
    SysUtils,
    Classes,
  {$ENDIF}
  {$IFNDEF VCL6}
    DsgnIntf,
  {$ELSE}
    DesignIntf,
    DesignEditors,
  {$ENDIF}
  QImport3StrTypes;

type
// ********************
// Property editors ***
// ********************
  TQFileNameProperty = class(TStringProperty)
  protected
    function GetDefaultExt: string; virtual; abstract;
    function GetFilter: string; virtual; abstract;
  public
    procedure Edit; override;
    function GetAttributes: TPropertyAttributes; override;
  end;

  TQASCIIFileNameProperty = class(TQFileNameProperty)
  protected
    function GetDefaultExt: string; override;
    function GetFilter: string; override;
  end;

  TQDBFFileNameProperty = class(TQFileNameProperty)
  protected
    function GetDefaultExt: string; override;
    function GetFilter: string; override;
  end;

  TQXMLFileNameProperty = class(TQFileNameProperty)
  protected
    function GetDefaultExt: string; override;
    function GetFilter: string; override;
  end;

  TQXLSFileNameProperty = class(TQFileNameProperty)
  protected
    function GetDefaultExt: string; override;
    function GetFilter: string; override;
  end;

  TQTemplateFileNameProperty = class(TQFileNameProperty)
  protected
    function GetDefaultExt: string; override;
    function GetFilter: string; override;
  end;

  TQWizardFileNameProperty = class(TQFileNameProperty)
  protected
    function GetDefaultExt: string; override;
    function GetFilter: string; override;
  end;

  TQImportAboutProperty = class(TStringProperty)
  public
    function GetAttributes: TPropertyAttributes; override;
    function GetValue: string; override;
    procedure Edit; override;
  end;

  TQImportVersionProperty = class(TStringProperty)
  public
    function GetAttributes: TPropertyAttributes; override;
    function GetValue: string; override;
  end;

  TQFieldNameProperty = class(TStringProperty)
  public
    function GetAttributes: TPropertyAttributes; override;
    procedure GetValues(Proc: TGetStrProc); override;
  end;

{$IFDEF HTML}
{$IFDEF VCL6}
  TQHTMLFileNameProperty = class(TQFileNameProperty)
  protected
    function GetDefaultExt: string; override;
    function GetFilter: string; override;
  end;
{$ENDIF}
{$ENDIF}

{$IFDEF XMLDOC}
{$IFDEF VCL6}
  TQXMLDocFileNameProperty = class(TQFileNameProperty)
  protected
    function GetDefaultExt: string; override;
    function GetFilter: string; override;
  end;
{$ENDIF}
{$ENDIF}

{$IFDEF XLSX}
{$IFDEF VCL6}
  TQXlsxFileNameProperty = class(TQFileNameProperty)
  protected
    function GetDefaultExt: string; override;
    function GetFilter: string; override;
  end;
{$ENDIF}
{$ENDIF}

{$IFDEF DOCX}
{$IFDEF VCL6}
  TQDocxFileNameProperty = class(TQFileNameProperty)
  protected
    function GetDefaultExt: string; override;
    function GetFilter: string; override;
  end;
{$ENDIF}
{$ENDIF}

{$IFDEF ODS}
{$IFDEF VCL6}
  TQODSFileNameProperty = class(TQFileNameProperty)
  protected
    function GetDefaultExt: string; override;
    function GetFilter: string; override;
  end;
{$ENDIF}
{$ENDIF}

{$IFDEF ODT}
{$IFDEF VCL6}
  TQODTFileNameProperty = class(TQFileNameProperty)
  protected
    function GetDefaultExt: string; override;
    function GetFilter: string; override;
  end;
{$ENDIF}
{$ENDIF}

{$IFDEF ADO}
  TQAccessFileNameProperty = class(TQFileNameProperty)
  protected
    function GetDefaultExt: string; override;
    function GetFilter: string; override;
  end;

  TQAccessTableNameProperty = class(TStringProperty)
  public
    function GetAttributes: TPropertyAttributes; override;
    procedure GetValues(Proc: TGetStrProc); override;
  end;
{$ENDIF}

// *************************
// *** Component editors ***
// *************************

  TQImportWizardEditor = class(TComponentEditor)
    function GetVerbCount: integer; override;
    function GetVerb(Index: integer): string; override;
    procedure ExecuteVerb(Index: integer); override;
  end;

  TQImportXLSEditor = class(TComponentEditor)
    function GetVerbCount: integer; override;
    function GetVerb(Index: integer): string; override;
    procedure ExecuteVerb(Index: integer); override;
  end;

  TQImportDataSetEditor = class(TComponentEditor)
    function GetVerbCount: integer; override;
    function GetVerb(Index: integer): string; override;
    procedure ExecuteVerb(Index: integer); override;
  end;

  TQImportXMLEditor = class(TComponentEditor)
    function GetVerbCount: integer; override;
    function GetVerb(Index: integer): string; override;
    procedure ExecuteVerb(Index: integer); override;
  end;

  TQImportDBFEditor = class(TComponentEditor)
    function GetVerbCount: integer; override;
    function GetVerb(Index: integer): string; override;
    procedure ExecuteVerb(Index: integer); override;
  end;

  TQImportASCIIEditor = class(TComponentEditor)
    function GetVerbCount: integer; override;
    function GetVerb(Index: integer): string; override;
    procedure ExecuteVerb(Index: integer); override;
  end;

{$IFDEF HTML}
{$IFDEF VCL6}
  TQImportHTMLEditor = class(TComponentEditor)
    function GetVerbCount: integer; override;
    function GetVerb(Index: integer): string; override;
    procedure ExecuteVerb(Index: integer); override;
  end;
{$ENDIF}
{$ENDIF}

{$IFDEF XMLDOC}
{$IFDEF VCL6}
  TQImportXMLDocEditor = class(TComponentEditor)
    function GetVerbCount: integer; override;
    function GetVerb(Index: integer): string; override;
    procedure ExecuteVerb(Index: integer); override;
  end;
{$ENDIF}
{$ENDIF}

{$IFDEF XLSX}
{$IFDEF VCL6}
  TQImportXlsxEditor = class(TComponentEditor)
    function GetVerbCount: integer; override;
    function GetVerb(Index: integer): string; override;
    procedure ExecuteVerb(Index: integer); override;
  end;
{$ENDIF}
{$ENDIF}

{$IFDEF DOCX}
{$IFDEF VCL6}
  TQImportDocxEditor = class(TComponentEditor)
    function GetVerbCount: integer; override;
    function GetVerb(Index: integer): string; override;
    procedure ExecuteVerb(Index: integer); override;
  end;
{$ENDIF}
{$ENDIF}

{$IFDEF ODS}
{$IFDEF VCL6}
  TQImportODSEditor = class(TComponentEditor)
    function GetVerbCount: integer; override;
    function GetVerb(Index: integer): string; override;
    procedure ExecuteVerb(Index: integer); override;
  end;
{$ENDIF}
{$ENDIF}

{$IFDEF ODT}
{$IFDEF VCL6}
  TQImportODTEditor = class(TComponentEditor)
    function GetVerbCount: integer; override;
    function GetVerb(Index: integer): string; override;
    procedure ExecuteVerb(Index: integer); override;
  end;
{$ENDIF}
{$ENDIF}

procedure Register;

implementation

uses
  {$IFDEF VCL16}
    Data.DB,
    Vcl.Dialogs,
  {$ELSE}
    DB,
    Dialogs,
  {$ENDIF}
  QImport3,
  QImport3XLS,
  QImport3DBF,
  QImport3ASCII,
  QImport3HTML,
  QImport3XML,
  QImport3DataSet,
  QImport3Wizard,
  QImport3Xlsx,
  QImport3Docx,
  QImport3ODS,
  QImport3ODT,
  {$IFDEF USESCRIPT}
  QImport3JScriptEngine,
  {$ENDIF}
  {$IFDEF ADO}
    ADO_QImport3Access,
  {$ENDIF}
  QImport3Common,
  QImport3StrIDs,
  fuQImport3DBFEditor,
  fuQImport3XMLEditor,
  fuQImport3TXTEditor,
  fuQImport3CSVEditor,
  fuQImport3DataSetEditor,
  fuQImport3About,
  fuQImport3FormatsEditor,
  fuQImport3XLSEditor,
  QImport3XMLDoc,
  fuQImport3XMLDocEditor,
  fuQImport3HTMLEditor,
  fuQImport3XlsxEditor,
  fuQImport3ODSEditor,
  fuQImport3ODTEditor,
  fuQImport3DocxEditor;


{$IFDEF VCL10}
{$R QImport3RegD10.dcr}
{$ELSE}
{$R QImport3Reg.dcr}
{$ENDIF}

procedure Register;
begin
  RegisterComponents(QI_PALETTE_PAGE,
    [TQImport3Wizard, TQImport3XLS, TQImport3DBF, TQImport3ASCII, TQImport3XML,
     TQImport3DataSet
    {$IFDEF USESCRIPT}, TQImport3JScriptEngine{$ENDIF}
    {$IFDEF HTML}{$IFDEF VCL6}, TQImport3HTML {$ENDIF}{$ENDIF}
    {$IFDEF XMLDOC}{$IFDEF VCL6}, TQImport3XMLDoc {$ENDIF}{$ENDIF}
    {$IFDEF DOCX}{$IFDEF VCL6}, TQImport3Xlsx {$ENDIF}{$ENDIF}
    {$IFDEF XLSX}{$IFDEF VCL6}, TQImport3Docx {$ENDIF}{$ENDIF}
    {$IFDEF ODS}{$IFDEF VCL6}, TQImport3ODS {$ENDIF}{$ENDIF}
    {$IFDEF ODT}{$IFDEF VCL6}, TQImport3ODT {$ENDIF}{$ENDIF}
    {$IFDEF ADO}, TADO_QImport3Access {$ENDIF}]);

// *********************************
// *** Register property editors ***
// *********************************

  RegisterPropertyEditor(TypeInfo(string), TQImport3ASCII, 'FileName',
                         TQASCIIFileNameProperty);

  RegisterPropertyEditor(TypeInfo(string), TQImport3DBF, 'FileName',
                         TQDBFFileNameProperty);

  RegisterPropertyEditor(TypeInfo(string), TQImport3XLS, 'FileName',
                         TQXLSFileNameProperty);

  RegisterPropertyEditor(TypeInfo(string), TQImport3Wizard, 'FileName',
                         TQWizardFileNameProperty);

  RegisterPropertyEditor(TypeInfo(string), TQImport3, 'About',
                         TQImportAboutProperty);
  RegisterPropertyEditor(TypeInfo(string), TQImport3, 'Version',
                         TQImportVersionProperty);

  RegisterPropertyEditor(TypeInfo(string), TQImport3Wizard, 'About',
                         TQImportAboutProperty);
  RegisterPropertyEditor(TypeInfo(string), TQImport3Wizard, 'Version',
                         TQImportVersionProperty);
  RegisterPropertyEditor(TypeInfo(string), TQImport3Wizard, 'TemplateFileName',
                         TQTemplateFileNameProperty);

  RegisterPropertyEditor(TypeInfo(string), TQImportFieldFormat, 'FieldName',
                         TQFieldNameProperty);
{$IFDEF HTML}
{$IFDEF VCL6}
  RegisterPropertyEditor(TypeInfo(string), TQImport3HTML, 'FileName',
                         TQHTMLFileNameProperty);
{$ENDIF}
{$ENDIF}
  RegisterPropertyEditor(TypeInfo(string), TQImport3XML, 'FileName',
                         TQXMLFileNameProperty);
{$IFDEF XMLDOC}
{$IFDEF VCL6}
  RegisterPropertyEditor(TypeInfo(string), TQImport3XMLDoc, 'FileName',
                         TQXMLDocFileNameProperty);
{$ENDIF}
{$ENDIF}
{$IFDEF XLSX}
{$IFDEF VCL6}
  RegisterPropertyEditor(TypeInfo(string), TQImport3Xlsx, 'FileName',
                         TQXlsxFileNameProperty);
{$ENDIF}
{$ENDIF}
{$IFDEF DOCX}
{$IFDEF VCL6}
  RegisterPropertyEditor(TypeInfo(string), TQImport3Docx, 'FileName',
                         TQDocxFileNameProperty);
{$ENDIF}
{$ENDIF}
{$IFDEF ODS}
{$IFDEF VCL6}
  RegisterPropertyEditor(TypeInfo(string), TQImport3ODS, 'FileName',
                         TQODSFileNameProperty);
{$ENDIF}
{$ENDIF}
{$IFDEF ODT}
{$IFDEF VCL6}
  RegisterPropertyEditor(TypeInfo(string), TQImport3ODT, 'FileName',
                         TQODTFileNameProperty);
{$ENDIF}
{$ENDIF}
{$IFDEF ADO}
  RegisterPropertyEditor(TypeInfo(string), TADO_QImport3Access, 'FileName',
                         TQAccessFileNameProperty);
  RegisterPropertyEditor(TypeInfo(string), TADO_QImport3Access, 'TableName',
                         TQAccessTableNameProperty);
{$ENDIF}

// *********************************
// *** Register component editors ***
// *********************************

  RegisterComponentEditor(TQImport3Wizard, TQImportWizardEditor);
  RegisterComponentEditor(TQImport3XLS, TQImportXLSEditor);
  RegisterComponentEditor(TQImport3DataSet, TQImportDataSetEditor);
  RegisterComponentEditor(TQImport3DBF, TQImportDBFEditor);
  RegisterComponentEditor(TQImport3XML, TQImportXMLEditor);
  RegisterComponentEditor(TQImport3ASCII, TQImportASCIIEditor);
{$IFDEF HTML}
{$IFDEF VCL6}
  RegisterComponentEditor(TQImport3HTML, TQImportHTMLEditor);
{$ENDIF}
{$ENDIF}
{$IFDEF XMLDOC}
{$IFDEF VCL6}
  RegisterComponentEditor(TQImport3XMLDoc, TQImportXMLDocEditor);
{$ENDIF}
{$ENDIF}
{$IFDEF XLSX}
{$IFDEF VCL6}
  RegisterComponentEditor(TQImport3Xlsx, TQImportXlsxEditor);
{$ENDIF}
{$ENDIF}
{$IFDEF DOCX}
{$IFDEF VCL6}
  RegisterComponentEditor(TQImport3Docx, TQImportDocxEditor);
{$ENDIF}
{$ENDIF}
{$IFDEF ODS}
{$IFDEF VCL6}
  RegisterComponentEditor(TQImport3ODS, TQImportODSEditor);
{$ENDIF}
{$ENDIF}
{$IFDEF ODT}
{$IFDEF VCL6}
  RegisterComponentEditor(TQImport3ODT, TQImportODTEditor);
{$ENDIF}
{$ENDIF}
end;

{ TQFileNameProperty }

procedure TQFileNameProperty.Edit;
var
  OD: TOpenDialog;
begin
  OD := TOpenDialog.Create(nil);
  try
    OD.DefaultExt := GetDefaultExt;
    OD.Filter := GetFilter;
    OD.FileName := GetStrValue;
    if OD.Execute then SetStrValue(OD.FileName);
  finally
    OD.Free;
  end;
end;

function TQFileNameProperty.GetAttributes: TPropertyAttributes;
begin
  Result := inherited GetAttributes + [paDialog];
end;

{ TQASCIIFileNameProperty }

function TQASCIIFileNameProperty.GetDefaultExt: string;
begin
  Result := 'txt';
end;

function TQASCIIFileNameProperty.GetFilter: string;
begin
  Result := QImportLoadStr(QIF_ASCII);
end;

{ TQDBFFileNameProperty }

function TQDBFFileNameProperty.GetDefaultExt: string;
begin
  Result := 'dbf';
end;

function TQDBFFileNameProperty.GetFilter: string;
begin
  Result := QImportLoadStr(QIF_DBF);
end;

{ TQHTMLFileNameProperty }

{$IFDEF HTML}
{$IFDEF VCL6}
function TQHTMLFileNameProperty.GetDefaultExt: string;
begin
  Result := 'html';
end;

function TQHTMLFileNameProperty.GetFilter: string;
begin
  Result := 'HTML files (*.htm, *.html)|*.htm; *.html';
end;
{$ENDIF}
{$ENDIF}

{ TQXMLFileNameProperty }

function TQXMLFileNameProperty.GetDefaultExt: string;
begin
  Result := 'xml';
end;

function TQXMLFileNameProperty.GetFilter: string;
begin
  Result := QImportLoadStr(QIF_XML);
end;

{ TQXMLDocFileNameProperty }

{$IFDEF XMLDOC}
{$IFDEF VCL6}
function TQXMLDocFileNameProperty.GetDefaultExt: string;
begin
  Result := 'xml';
end;

function TQXMLDocFileNameProperty.GetFilter: string;
begin
  Result := QImportLoadStr(QIF_XML);
end;
{$ENDIF}
{$ENDIF}

{ TQXLSFileNameProperty }

function TQXLSFileNameProperty.GetDefaultExt: string;
begin
  Result := 'xls';
end;

function TQXLSFileNameProperty.GetFilter: string;
begin
  Result := QImportLoadStr(QIF_XLS);
end;

{ TQTemplateFileNameProperty }

function TQTemplateFileNameProperty.GetDefaultExt: string;
begin
  Result := 'imp';
end;

function TQTemplateFileNameProperty.GetFilter: string;
begin
  Result := QImportLoadStr(QIF_TEMPLATE);
end;

{ TQWizardFileNameProperty }

function TQWizardFileNameProperty.GetDefaultExt: string;
begin
  Result := EmptyStr;
end;

function TQWizardFileNameProperty.GetFilter: string;
begin
  Result := QImportLoadStr(QIF_ALL);
end;

{ TQImportAboutProperty }

procedure TQImportAboutProperty.Edit;
begin
  ShowAboutForm;
end;

function TQImportAboutProperty.GetAttributes: TPropertyAttributes;
begin
  Result := [paReadOnly, paDialog];
end;

function TQImportAboutProperty.GetValue: string;
begin
  Result := QI_ABOUT;
end;

{ TQImportVersionProperty }

function TQImportVersionProperty.GetAttributes: TPropertyAttributes;
begin
  Result := [paReadOnly];
end;

function TQImportVersionProperty.GetValue: string;
begin
  Result := QI_VERSION;
end;

{ TQImportDBFEditor }

function TQImportDBFEditor.GetVerbCount: integer;
begin
  Result := 2;
end;

function TQImportDBFEditor.GetVerb(Index: integer): string;
begin
  case Index of
    0: Result := 'Map...';
    1: Result := 'Field formats...';
  end;
end;

procedure TQImportDBFEditor.ExecuteVerb(Index: integer);
begin
  case Index of
    0: RunQImportDBFEditor(Component as TQImport3DBF);
    1: if RunFormatsEditor(Component as TQImport3) then Designer.Modified;
  end;
end;

{ TQImportHTMLEditor }

{$IFDEF HTML}
{$IFDEF VCL6}
function TQImportHTMLEditor.GetVerbCount: integer;
begin
  Result := 2;
end;

function TQImportHTMLEditor.GetVerb(Index: integer): string;
begin
  case Index of
    0: Result := 'Map...';
    1: Result := 'Field formats...';
  end;
end;

procedure TQImportHTMLEditor.ExecuteVerb(Index: integer);
var
  NeedModify: boolean;
begin
  NeedModify := false;
  case Index of
    0: RunQImportHTMLEditor(Component as TQImport3HTML);
    1: NeedModify := RunFormatsEditor(Component as TQImport3);
  end;
  if NeedModify then Designer.Modified;
end;
{$ENDIF}
{$ENDIF}

{ TQImportXMLEditor }

function TQImportXMLEditor.GetVerbCount: integer;
begin
  Result := 2;
end;

function TQImportXMLEditor.GetVerb(Index: integer): string;
begin
  case Index of
    0: Result := 'Map...';
    1: Result := 'Field formats...';
  end;
end;

procedure TQImportXMLEditor.ExecuteVerb(Index: integer);
begin
  case Index of
    0: RunQImportXMLEditor(Component as TQImport3XML);
    1: if RunFormatsEditor(Component as TQImport3) then Designer.Modified;
  end;
end;

{ TQImportDataSetEditor }

function TQImportDataSetEditor.GetVerbCount: integer;
begin
  Result := 2;
end;

function TQImportDataSetEditor.GetVerb(Index: integer): string;
begin
  case Index of
    0: Result := 'Map...';
    1: Result := 'Field formats...';
  end;
end;

procedure TQImportDataSetEditor.ExecuteVerb(Index: integer);
begin
  case Index of
    0: RunQImportDataSetEditor(Component as TQImport3DataSet);
    1: if RunFormatsEditor(Component as TQImport3) then Designer.Modified;
  end;
end;

{ TQImportASCIIEditor }

function TQImportASCIIEditor.GetVerbCount: integer;
begin
  Result := 2;
end;

function TQImportASCIIEditor.GetVerb(Index: integer): string;
begin
  case Index of
    0: Result := 'Map...';
    1: Result := 'Field formats...';
  end;
end;

procedure TQImportASCIIEditor.ExecuteVerb(Index: integer);
var
  NeedModify: boolean;
begin
  NeedModify := false;
  case Index of
    0: if Component is TQImport3ASCII then
       begin
         if (Component as TQImport3ASCII).Comma = #0
           then NeedModify := RunQImportTXTEditor(Component as TQImport3ASCII)
           else NeedModify := RunQImportCSVEditor(Component as TQImport3ASCII);
       end;
    1: NeedModify := RunFormatsEditor(Component as TQImport3);
  end;
  if NeedModify then
    Designer.Modified;
end;

{ TQImportXLSEditor }

function TQImportXLSEditor.GetVerbCount: integer;
begin
  Result := 2;
end;

function TQImportXLSEditor.GetVerb(Index: integer): string;
begin
  case Index of
    0: Result := 'Map...';
    1: Result := 'Field formats...';
  end;
end;

procedure TQImportXLSEditor.ExecuteVerb(Index: integer);
var
  NeedModify: boolean;
begin
  NeedModify := false;
  case Index of
    0: NeedModify := RunQImportXLSEditor(Component as TQImport3XLS);
    1: NeedModify := RunFormatsEditor(Component as TQImport3);
  end;
  if NeedModify then Designer.Modified;
end;

{ TQFieldNameProperty }

function TQFieldNameProperty.GetAttributes: TPropertyAttributes;
begin
  Result := [paValueList]
end;

procedure TQFieldNameProperty.GetValues(Proc: TGetStrProc);
var
  WasInActive: Boolean;
  Import: TQImport3;
  Wizard: TQImport3Wizard;
  i: Integer;
  Component: TComponent;
  DataSet: TDataSet;
begin
  Import := nil;
  Wizard := nil;

  Component := (((GetComponent(0) as TQImportFieldFormat).Collection as
    TQImportFieldFormats).Holder);
  if Component is TQImport3 then
    Import := Component as TQImport3
  else if Component is TQImport3Wizard then
    Wizard := Component as TQImport3Wizard;

  if Assigned(Import) then
    DataSet := Import.DataSet
  else if Assigned(Wizard) then
    DataSet := Wizard.DataSet
  else
    DataSet := nil;

  if not Assigned(DataSet) then Exit;

  WasInActive :=  not Dataset.Active;
  if WasInActive then
  try
    Dataset.Open;
  except
    Exit;
  end;
  for I := 0 to DataSet.FieldCount - 1 do
  begin
    Proc(DataSet.Fields[I].FieldName);
  end;
  if WasInActive then
  try
    Dataset.Close;
  except
  end;
end;

{ TQImportWizardEditor }

procedure TQImportWizardEditor.ExecuteVerb(Index: integer);
begin
  if RunFormatsEditor(Component as TComponent{TQImport3Wizard}) then
    Designer.Modified;
end;

function TQImportWizardEditor.GetVerb(Index: integer): string;
begin
  Result := 'Field formats...'
end;

function TQImportWizardEditor.GetVerbCount: integer;
begin
  Result := 1;
end;

{ TQImportXMLDocEditor }

{$IFDEF XMLDOC}
{$IFDEF VCL6}
function TQImportXMLDocEditor.GetVerbCount: integer;
begin
  Result := 2;
end;

function TQImportXMLDocEditor.GetVerb(Index: integer): string;
begin
  case Index of
    0: Result := 'Map...';
    1: Result := 'Field formats...';
  end;
end;

procedure TQImportXMLDocEditor.ExecuteVerb(Index: integer);
begin
  case Index of
    0: RunQImportXMLDocEditor(Component as TQImport3XMLDoc);
    1: if RunFormatsEditor(Component as TQImport3) then Designer.Modified;
  end;
end;
{$ENDIF}
{$ENDIF}

{$IFDEF ADO}
function TQAccessFileNameProperty.GetDefaultExt: string;
begin
  Result := 'mdb';
end;

function TQAccessFileNameProperty.GetFilter: string;
begin
  Result := QImportLoadStr(QIF_Access);
end;

{ TQAccessTableNameProperty }

function TQAccessTableNameProperty.GetAttributes: TPropertyAttributes;
begin
  Result := [paValueList, paSortList];
end;

procedure TQAccessTableNameProperty.GetValues(Proc: TGetStrProc);
var
  Imp: TADO_QImport3Access;
  List: TStrings;
  i: integer;
begin
  Imp := GetComponent(0) as TADO_QImport3Access;
  List := TStringList.Create;
  try
    Imp.GetTableNames(List);
    for i := 0 to List.Count - 1 do Proc(List[i]);
  finally
    List.Free;
  end;
end;
{$ENDIF}

{ TQXlsxFileNameProperty }

{$IFDEF XLSX}
{$IFDEF VCL6}
function TQXlsxFileNameProperty.GetDefaultExt: string;
begin
  Result := 'xlsx';
end;

function TQXlsxFileNameProperty.GetFilter: string;
begin
  Result := 'Microsoft Excel 2007 files (*.xlsx)|*.xlsx';
end;
{$ENDIF}
{$ENDIF}

{ TQDocxFileNameProperty }

{$IFDEF DOCX}
{$IFDEF VCL6}
function TQDocxFileNameProperty.GetDefaultExt: string;
begin
  Result := 'docx';
end;

function TQDocxFileNameProperty.GetFilter: string;
begin
  Result := 'Microsoft Word 2007 files (*.docx)|*.docx';
end;
{$ENDIF}
{$ENDIF}

{ TQODSFileNameProperty }

{$IFDEF ODS}
{$IFDEF VCL6}
function TQODSFileNameProperty.GetDefaultExt: string;
begin
  Result := 'ods';
end;

function TQODSFileNameProperty.GetFilter: string;
begin
  Result := 'Open Document Spreadsheet files (*.ods)|*.ods';
end;
{$ENDIF}
{$ENDIF}

{ TQODTFileNameProperty }

{$IFDEF ODT}
{$IFDEF VCL6}
function TQODTFileNameProperty.GetDefaultExt: string;
begin
  Result := 'odt';
end;

function TQODTFileNameProperty.GetFilter: string;
begin
  Result := 'Open Document Text files (*.odt)|*.odt';
end;
{$ENDIF}
{$ENDIF}

{ TQImportXlsxEditor }

{$IFDEF XLSX}
{$IFDEF VCL6}
procedure TQImportXlsxEditor.ExecuteVerb(Index: integer);
var
  NeedModify: boolean;
begin
  NeedModify := False;
  case Index of
    0: RunQImportXlsxEditor(Component as TQImport3Xlsx);
    1: NeedModify := RunFormatsEditor(Component as TQImport3);
  end;
  if NeedModify then Designer.Modified;
end;

function TQImportXlsxEditor.GetVerb(Index: integer): string;
begin
  case Index of
    0: Result := 'Map...';
    1: Result := 'Field formats...';
  end;
end;

function TQImportXlsxEditor.GetVerbCount: integer;
begin
  Result := 2;
end;
{$ENDIF}
{$ENDIF}

{ TQImportDocxEditor }

{$IFDEF DOCX}
{$IFDEF VCL6}
procedure TQImportDocxEditor.ExecuteVerb(Index: integer);
var
  NeedModify: boolean;
begin
  NeedModify := False;
  case Index of
    0: RunQImportDocxEditor(Component as TQImport3Docx);
    1: NeedModify := RunFormatsEditor(Component as TQImport3);
  end;
  if NeedModify then Designer.Modified;
end;

function TQImportDocxEditor.GetVerb(Index: integer): string;
begin
  case Index of
    0: Result := 'Map...';
    1: Result := 'Field formats...';
  end;
end;

function TQImportDocxEditor.GetVerbCount: integer;
begin
  Result := 2;
end;
{$ENDIF}
{$ENDIF}

{ TQImportODSEditor }

{$IFDEF ODS}
{$IFDEF VCL6}
procedure TQImportODSEditor.ExecuteVerb(Index: integer);
var
  NeedModify: boolean;
begin
  NeedModify := False;
  case Index of
    0: RunQImportODSEditor(Component as TQImport3ODS);
    1: NeedModify := RunFormatsEditor(Component as TQImport3);
  end;
  if NeedModify then Designer.Modified;
end;

function TQImportODSEditor.GetVerb(Index: integer): string;
begin
  case Index of
    0: Result := 'Map...';
    1: Result := 'Field formats...';
  end;
end;

function TQImportODSEditor.GetVerbCount: integer;
begin
  Result := 2;
end;
{$ENDIF}
{$ENDIF}

{ TQImportODTEditor }

{$IFDEF ODT}
{$IFDEF VCL6}
procedure TQImportODTEditor.ExecuteVerb(Index: integer);
var
  NeedModify: boolean;
begin
  NeedModify := False;
  case Index of
    0: RunQImportODTEditor(Component as TQImport3ODT);
    1: NeedModify := RunFormatsEditor(Component as TQImport3);
  end;
  if NeedModify then Designer.Modified;
end;

function TQImportODTEditor.GetVerb(Index: integer): string;
begin
  case Index of
    0: Result := 'Map...';
    1: Result := 'Field formats...';
  end;
end;

function TQImportODTEditor.GetVerbCount: integer;
begin
  Result := 2;
end;
{$ENDIF}
{$ENDIF}

end.
