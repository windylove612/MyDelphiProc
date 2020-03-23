unit QImport3Wizard;

{$I QImport3VerCtrl.Inc}

interface
                          
uses
  {$IFDEF VCL16}
    Vcl.Forms,
    Vcl.Dialogs,
    Winapi.Windows,
    Vcl.Grids,
    Vcl.ComCtrls,
    Vcl.Buttons,
    Vcl.Controls,
    Vcl.ExtCtrls,
    System.Classes,
    Data.Db,
    Vcl.ToolWin,
    Vcl.DBGrids,
    Vcl.StdCtrls,
    Vcl.ImgList,
    Vcl.Menus,
    Vcl.Graphics,
  {$ELSE}
    Forms,
    Dialogs,
    Windows,
    Grids,
    ComCtrls,
    Buttons,
    Controls,
    ExtCtrls,
    Classes,
    Db,
    ToolWin,
    DBGrids,
    StdCtrls,
    ImgList,
    Menus,
    Graphics,
  {$ENDIF}
  QImport3XLS,
  QImport3ASCII,
  QImport3,
  QImport3DBF,
  fuQImport3ProgressDlg,
  QImport3TXTView,
  QImport3XLSFile,
  QImport3XLSMapParser,
  QImport3InfoPanel,
  QImport3HTML,
  QImport3XML,
  {$IFDEF XMLDOC}QImport3XMLDoc,{$ENDIF}
  {$IFDEF XLSX}QImport3Xlsx,{$ENDIF}
  {$IFDEF DOCX}QImport3Docx,{$ENDIF}
  {$IFDEF ODS}QImport3ODS,{$ENDIF}
  {$IFDEF ODT}QImport3ODT,{$ENDIF}
  {$IFDEF ADO}ADO_QImport3Access,{$ENDIF}
  {$IFDEF USESCRIPT}QImport3JScriptEngine, QImport3ScriptEngine,{$ENDIF}
  QImport3StrTypes;

type

  TQImport3Wizard = class(TComponent)
  private
    FAllowedImports: TAllowedImports;

    FDataSet: TDataSet;
    FDBGrid: TDBGrid;
    FListView: TListView;
    FStringGrid: TStringGrid;

    FFileName: qiString;
    FFormats: TQImportFormats;
    FFieldFormats: TQImportFieldFormats;
    FAbout: string;
    FVersion: string;

    FImportRecCount: Integer;
    FCommitRecCount: Integer;
    FCommitAfterDone: boolean;
    FErrorLog: boolean;
    FErrorLogFileName: qiString;
    FRewriteErrorLogFile: boolean;
    FShowErrorLog: boolean;
    FShowProgress: boolean;
    FAutoChangeExtension: boolean;
    FShowHelpButton: boolean;
    FCloseAfterImport: boolean;
    FPicture: TPicture;
    FTextViewerRows: Integer;
    FCSVViewerRows: Integer;
    FExcelViewerRows: Integer;
    FExcelMaxColWidth: Integer;
    FAutoTrimValue: Boolean;
    FImportEmptyRows: Boolean;

    FShowSaveLoadButtons: Boolean;

    FAutoSaveTemplate: Boolean;
    FAutoLoadTemplate: Boolean;
    FTemplateFileName: qiString;
    FGoToLastPage: boolean;

    FImportDestination: TQImportDestination;
    FImportMode: TQImportMode;
    FAddType: TQImportAddType;
    FKeyColumns: TStrings;
    FGridCaptionRow: Integer;
    FGridStartRow: Integer;
    FConfirmOnCancel: boolean;

    FOnBeforeImport: TNotifyEvent;
    FOnAfterImport: TNotifyEvent;
    FOnImportRecord: TNotifyEvent;
    FOnImportError: TNotifyEvent;
    FOnImportErrorAdv: TNotifyEvent;
    FOnNeedCommit: TNotifyEvent;
    FOnImportCancel: TImportCancelEvent;
    FOnBeforePost: TImportBeforePostEvent;
    FOnLoadTemplate: TImportLoadTemplateEvent;
    FOnDestinationLocate: TDestinationLocateEvent;

    function IsFileName: Boolean;
    procedure SetFormats(const Value: TQImportFormats);
    procedure SetFieldFormats(const Value: TQImportFieldFormats);
    procedure SetKeyColumns(const Value: TStrings);
    procedure SetPicture(const Value: TPicture);
  protected
    procedure Notification(AComponent: TComponent;
      Operation: TOperation); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Execute;
  published
    property AllowedImports: TAllowedImports read FAllowedImports
      write FAllowedImports default [Low(TAllowedImport)..High(TAllowedImport)];

    property DataSet: TDataSet read FDataSet write FDataSet;
    property DBGrid: TDBGrid read FDBGrid write FDBGrid;
    property ListView: TListView read FListView write FListView;
    property StringGrid: TStringGrid read FStringGrid write FStringGrid;

    property FileName: qiString read FFileName write FFileName stored IsFileName;
    property Formats: TQImportFormats read FFormats write SetFormats;
    property FieldFormats: TQImportFieldFormats read FFieldFormats
      write SetFieldFormats;
    property About: string read FAbout write FAbout;
    property Version: string read FVersion write FVersion;

    property ImportRecCount: Integer read FImportRecCount
      write FImportRecCount default 0;
    property AutoTrimValue: Boolean read FAutoTrimValue
      write FAutoTrimValue default False;
    property ImportEmptyRows: Boolean read FImportEmptyRows
      write FImportEmptyRows default True;

    property CommitRecCount: Integer read FCommitRecCount
      write FCommitRecCount default 1000;
    property CommitAfterDone: boolean read FCommitAfterDone
      write FCommitAfterDone default False;
    property ErrorLog: boolean read FErrorLog write FErrorLog default False;
    property ErrorLogFileName: qiString read FErrorLogFileName
      write FErrorLogFileName;
    property RewriteErrorLogFile: boolean read FRewriteErrorLogFile
      write FRewriteErrorLogFile default True;
    property ShowErrorLog: boolean read FShowErrorLog
      write FShowErrorLog default False;
    property ShowProgress: boolean read FShowProgress
      write FShowProgress default True;
    property AutoChangeExtension: boolean read FAutoChangeExtension
      write FAutoChangeExtension default True;
    property ShowHelpButton: boolean read FShowHelpButton
      write FShowHelpButton default True;
    property CloseAfterImport: boolean read FCloseAfterImport
      write FCloseAfterImport default False;
    property Picture: TPicture read FPicture write SetPicture;
    property TextViewerRows: Integer read FTextViewerRows
      write FTextViewerRows default 20;
    property CSVViewerRows: Integer read FCSVViewerRows
      write FCSVViewerRows default 20;
    property ExcelViewerRows: Integer read FExcelViewerRows
      write FExcelViewerRows default 256;
    property ExcelMaxColWidth: Integer read FExcelMaxColWidth
      write FExcelMaxColWidth default 130;

    property ShowSaveLoadButtons: boolean read FShowSaveLoadButtons
      write FShowSaveLoadButtons default True;

    property TemplateFileName: qiString read FTemplateFileName
      write FTemplateFileName;
    property AutoLoadTemplate: boolean read FAutoLoadTemplate
      write FAutoLoadTemplate default False;
    property AutoSaveTemplate: boolean read FAutoSaveTemplate
      write FAutoSaveTemplate default False;
    property GoToLastPage: boolean read FGoToLastPage
      write FGoToLastPage default False;

    property ImportDestination: TQImportDestination read FImportDestination
      write FImportDestination default qidDataSet;
    property ImportMode: TQImportMode read FImportMode write FImportMode
      default qimInsertAll;
    property AddType: TQImportAddType read FAddType
      write FAddType default qatAppend;
    property KeyColumns: TStrings read FKeyColumns write SetKeyColumns;
    property GridCaptionRow: Integer read FGridCaptionRow
      write FGridCaptionRow default -1;
    property GridStartRow: Integer read FGridStartRow
      write FGridStartRow default -1;
    property ConfirmOnCancel: boolean read FConfirmOnCancel
      write FConfirmOnCancel default True;

    property OnBeforeImport: TNotifyEvent read FOnBeforeImport
      write FOnBeforeImport;
    property OnAfterImport: TNotifyEvent read FOnAfterImport
      write FOnAfterImport;
    property OnImportRecord: TNotifyEvent read FOnImportRecord
      write FOnImportRecord;
    property OnImportError: TNotifyEvent read FOnImportError
      write FOnImportError;
    property OnImportErrorAdv: TNotifyEvent read FOnImportErrorAdv
      write FOnImportErrorAdv;
    property OnNeedCommit: TNotifyEvent read FOnNeedCommit
      write FOnNeedCommit;
    property OnImportCancel: TImportCancelEvent read FOnImportCancel
      write FOnImportCancel;
    property OnBeforePost: TImportBeforePostEvent read FOnBeforePost
      write FOnBeforePost;
    property OnLoadTemplate: TImportLoadTemplateEvent read FOnLoadTemplate
      write FOnLoadTemplate;
    property OnDestinationLocate: TDestinationLocateEvent
      read FOnDestinationLocate write FOnDestinationLocate;
  end;

  TQImport3WizardF = class(TForm)
    paButtons: TPanel;
    Bevel1: TBevel;
    bHelp: TButton;
    bBack: TButton;
    bNext: TButton;
    bCancel: TButton;
    bOk: TButton;
    pgImport: TPageControl;
    tsImportType: TTabSheet;
    tsTXTOptions: TTabSheet;
    tsDBFOptions: TTabSheet;
    tsExcelOptions: TTabSheet;
    grpImportTypes: TGroupBox;
    laComma: TLabel;
    rbtXLS: TRadioButton;
    rbtDBF: TRadioButton;
    rbtTXT: TRadioButton;
    rbtCSV: TRadioButton;
    cbComma: TComboBox;
    laSourceFileName: TLabel;
    edtFileName: TEdit;
    Bevel6: TBevel;
    opnDialog: TOpenDialog;
    spbBrowse: TSpeedButton;
    tsCommitOptions: TTabSheet;
    laTXTStep_02: TLabel;
    Bevel2: TBevel;
    laTXTSkipLines: TLabel;
    edtTXTSkipLines: TEdit;
    laStep_03: TLabel;
    Bevel3: TBevel;
    lstDBFDataSet: TListView;
    lstDBF: TListView;
    lstDBFMap: TListView;
    bDBFAdd: TSpeedButton;
    bDBFRemove: TSpeedButton;
    laStep_04: TLabel;
    Bevel4: TBevel;
    laStep_02: TLabel;
    Bevel5: TBevel;
    tsFormats: TTabSheet;
    laStep_06: TLabel;
    Bevel7: TBevel;
    pgFormats: TPageControl;
    tshBaseFormats: TTabSheet;
    grpDateTimeFormats: TGroupBox;
    laShortDateFormat: TLabel;
    laLongDateFormat: TLabel;
    laShortTimeFormat: TLabel;
    laLongTimeFormat: TLabel;
    edtShortDateFormat: TEdit;
    edtLongDateFormat: TEdit;
    edtShortTimeFormat: TEdit;
    edtLongTimeFormat: TEdit;
    tshDataFormats: TTabSheet;
    lstFormatFields: TListView;
    ilWizard: TImageList;
    odTemplate: TOpenDialog;
    btnSaveTemplate: TSpeedButton;
    sdTemplate: TSaveDialog;
    bDBFAutoFill: TSpeedButton;
    bDBFClear: TSpeedButton;
    laQuote: TLabel;
    tsCSVOptions: TTabSheet;
    laStep_07: TLabel;
    Bevel15: TBevel;
    pgFieldOptions: TPageControl;
    tsFieldTuning: TTabSheet;
    laGeneratorValue: TLabel;
    edtGeneratorValue: TEdit;
    laGeneratorStep: TLabel;
    edtGeneratorStep: TEdit;
    laConstantValue: TLabel;
    edtConstantValue: TEdit;
    laNullValue: TLabel;
    edtNullValue: TEdit;
    laDefaultValue: TLabel;
    edtDefaultValue: TEdit;
    laLeftQuote: TLabel;
    edtLeftQuote: TEdit;
    laRightQuote: TLabel;
    edtRightQuote: TEdit;
    laQuoteAction: TLabel;
    cmbQuoteAction: TComboBox;
    laCharCase: TLabel;
    cmbCharCase: TComboBox;
    laCharSet: TLabel;
    cmbCharSet: TComboBox;
    Bevel13: TBevel;
    grpSeparators: TGroupBox;
    laDecimalSeparator: TLabel;
    edtDecimalSeparator: TEdit;
    laThousandSeparator: TLabel;
    edtThousandSeparator: TEdit;
    laDateSeparator: TLabel;
    edtDateSeparator: TEdit;
    laTimeSeparator: TLabel;
    edtTimeSeparator: TEdit;
    cbQuote: TComboBox;
    pcLastStep: TPageControl;
    tshCommit: TTabSheet;
    tshAdvanced: TTabSheet;
    grpImportCount: TGroupBox;
    laImportRecCount_01: TLabel;
    laImportRecCount_02: TLabel;
    chImportAllRecords: TCheckBox;
    edtImportRecCount: TEdit;
    grpCommit: TGroupBox;
    laCommitRecCount_01: TLabel;
    laCommitRecCount_02: TLabel;
    chCommitAfterDone: TCheckBox;
    edtCommitRecCount: TEdit;
    chDBFSkipDeleted: TCheckBox;
    sdErrorLog: TSaveDialog;
    chCloseAfterImport: TCheckBox;
    pcXLSFile: TPageControl;
    paXLSFieldsAndRanges: TPanel;
    lvXLSFields: TListView;
    lvXLSRanges: TListView;
    tbXLSRanges: TToolBar;
    tbtXLSAddRange: TToolButton;
    tbtXLSEditRange: TToolButton;
    tbtXLSDelRange: TToolButton;
    tbXLSUtils: TToolBar;
    tbtXLSAutoFillCols: TToolButton;
    tbtXLSAutoFillRows: TToolButton;
    tbtXLSMoveRangeUp: TToolButton;
    tbtXLSMoveRangeDown: TToolButton;
    tbtSeparator_01: TToolButton;
    tbtXLSClearFieldRanges: TToolButton;
    grpErrorLog: TGroupBox;
    laErrorLogFileName: TLabel;
    bvErrorLogFileName: TBevel;
    bErrorLogFileName: TSpeedButton;
    chEnableErrorLog: TCheckBox;
    chShowErrorLog: TCheckBox;
    edErrorLogFileName: TEdit;
    chRewriteErrorLogFile: TCheckBox;
    rgImportMode: TRadioGroup;
    lvAvailableColumns: TListView;
    laAvailableColumns: TLabel;
    bAllToRight: TSpeedButton;
    bOneToRirght: TSpeedButton;
    bOneToLeft: TSpeedButton;
    bAllToLeft: TSpeedButton;
    lvSelectedColumns: TListView;
    laSelectedColumns: TLabel;
    rgAddType: TRadioGroup;
    laReplacements: TLabel;
    gbTemplateOptions: TGroupBox;
    btnLoadTemplate: TSpeedButton;
    chGoToLastPage: TCheckBox;
    chAutoSaveTemplate: TCheckBox;
    pbDBFAdd: TPaintBox;
    pbDBFAutoFill: TPaintBox;
    pbDBFRemove: TPaintBox;
    pbDBFClear: TPaintBox;
    lvCSVFields: TListView;
    tbCSV: TToolBar;
    tbtCSVAutoFill: TToolButton;
    tbtCSVClear: TToolButton;
    laCSVSkipLines: TLabel;
    edtCSVSkipLines: TEdit;
    cbCSVColNumber: TComboBox;
    laCSVColNumber: TLabel;
    lvTXTFields: TListView;
    tbTXT: TToolBar;
    tbtTXTClear: TToolButton;
    laXLSSkipCols: TLabel;
    edXLSSkipCols: TEdit;
    laXLSSkipRows: TLabel;
    edXLSSkipRows: TEdit;
    rbtXML: TRadioButton;
    tsXMLOptions: TTabSheet;
    laStep_05: TLabel;
    Bevel8: TBevel;
    lvXMLDataSet: TListView;
    lvXML: TListView;
    lvXMLMap: TListView;
    bXMLAdd: TSpeedButton;
    pbXMLAdd: TPaintBox;
    bXMLAutoFill: TSpeedButton;
    pbXMLAutoFill: TPaintBox;
    bXMLRemove: TSpeedButton;
    pbXMLRemove: TPaintBox;
    bXMLClear: TSpeedButton;
    pbXMLClear: TPaintBox;
    mmBooleanTrue: TMemo;
    laBooleanTrue: TLabel;
    mmBooleanFalse: TMemo;
    laBooleanFalse: TLabel;
    laNullValues: TLabel;
    mmNullValues: TMemo;
    tbtXLSClearAllRanges: TToolButton;
    lvXLSSelection: TListView;
    lvReplacements: TListView;
    tbReplacements: TToolBar;
    tbtAddReplacement: TToolButton;
    tbtEditReplacement: TToolButton;
    tbtDelReplacement: TToolButton;
    chXMLWriteOnFly: TCheckBox;
    laTemplateFileName: TLabel;
    laTemplateFileNameTag: TLabel;
    tsHTMLOptions: TTabSheet;
    laStep_08: TLabel;
    Bevel9: TBevel;
    laHTMLSkipLines: TLabel;
    laHTMLColNumber: TLabel;
    lvHTMLFields: TListView;
    tbHTML: TToolBar;
    tbtHTMLAutoFill: TToolButton;
    tbtHTMLClear: TToolButton;
    edtHTMLSkipLines: TEdit;
    cbHTMLColNumber: TComboBox;
    laHTMLTableNumber: TLabel;
    cbHTMLTableNumber: TComboBox;
    rbtHTML: TRadioButton;
    Image: TImage;
    rbtXMLDoc: TRadioButton;
    tsXMLDocOptions: TTabSheet;
    laXMLDocColNumber: TLabel;
    laXMLDocXPath: TLabel;
    laStep_10: TLabel;
    Bevel10: TBevel;
    laXMLDocSkipLines: TLabel;
    lvXMLDocFields: TListView;
    edXMLDocSkipLines: TEdit;
    tbXMLDoc: TToolBar;
    tbtXMLDocAutoFill: TToolButton;
    tbtXMLDocClear: TToolButton;
    cbXMLDocColNumber: TComboBox;
    edXMLDocXPath: TEdit;
//    beDocxFillGrid: TBevel;
    bXMLDocFillGrid: TSpeedButton;
    tvXMLDoc: TTreeView;
    bXMLDocGetXPath: TSpeedButton;
    bXMLDocBuildTree: TSpeedButton;
    pmXMLDocTreeView: TPopupMenu;
    miGetXPath: TMenuItem;
    rbtAccess: TRadioButton;
    edAccessPassword: TEdit;
    laAccessPassword: TLabel;
    sdQuery: TSaveDialog;
    odQuery: TOpenDialog;
    tsAccessOptions_01: TTabSheet;
    tsAccessOptions_02: TTabSheet;
    Bevel12: TBevel;
    laStep_16: TLabel;
    rbtAccessTable: TRadioButton;
    rbtAccessSQL: TRadioButton;
    lbAccessTables: TListBox;
    memAccessSQL: TMemo;
    tbAccess: TToolBar;
    tbtAccessSQLLoad: TToolButton;
    tbtAccessSQLSave: TToolButton;
    Bevel16: TBevel;
    laStep_15: TLabel;
    bAccessAdd: TSpeedButton;
    pbAccessAdd: TPaintBox;
    bAccessAutoFill: TSpeedButton;
    pbAccessAutoFill: TPaintBox;
    bAccessRemove: TSpeedButton;
    pbAccessRemove: TPaintBox;
    bAccessClear: TSpeedButton;
    pbAccessClear: TPaintBox;
    lstAccessDataSet: TListView;
    lstAccess: TListView;
    lstAccessMap: TListView;
    rbtXlsx: TRadioButton;
    rbtDocx: TRadioButton;
    tsXlsxOptions: TTabSheet;
    edXlsxSkipRows: TEdit;
    paXlsxFields: TPanel;
    pcXlsxFile: TPageControl;
    tlbXlsxUtils: TToolBar;
    tlbXlsxAutoFill: TToolButton;
    tlbXlsxClear: TToolButton;
    laStep_11: TLabel;
    bvlTop: TBevel;
    laXlsxSkipRows: TLabel;
    tsDocxOptions: TTabSheet;
    bvl2: TBevel;
    laStep_12: TLabel;
    paDocxFields: TPanel;
    lvDocxFields: TListView;
    tlbDocxUtils: TToolBar;
    tlbDocxAutoFill: TToolButton;
    tlbDocxClear: TToolButton;
    edDocxSkipRows: TEdit;
    pcDocxFile: TPageControl;
    laDocxSkipRows: TLabel;
    lvXlsxFields: TListView;
    rbtODS: TRadioButton;
    tsODSOptions: TTabSheet;
    bvl14: TBevel;
    laStep_13: TLabel;
    lvODSFields: TListView;
    tlbODSUtils: TToolBar;
    tlbODSAutoFill: TToolButton;
    tlbODSClear: TToolButton;
    laODSSkipRows: TLabel;
    edODSSkipRows: TEdit;
    pcODSFile: TPageControl;
    impDBF: TQImport3DBF;
    impXLS: TQImport3XLS;
    impASCII: TQImport3ASCII;
    impXML: TQImport3XML;
    tsODTOptions: TTabSheet;
    bvl15: TBevel;
    laStep_14: TLabel;
    lvODTFields: TListView;
    laODTSkipRows: TLabel;
    tlbODTUtils: TToolBar;
    tlbODTAutoFill: TToolButton;
    tlbODTClear: TToolButton;
    pcODTFile: TPageControl;
    edODTSkipRows: TEdit;
    cbODTUseHeader: TCheckBox;
    rbtODT: TRadioButton;
    cbXMLDocDataLocation: TComboBox;
    laXMLDocDataLocation: TLabel;
    cmbTXTEncoding: TComboBox;
    laTXTEncoding: TLabel;
    cmbCSVEncoding: TComboBox;
    laCSVEncoding: TLabel;
    chImportEmptyRows: TCheckBox;
    chAutoTrimValue: TCheckBox;
    lScript: TLabel;
    mScript: TMemo;
    cbXMLDocumentType: TComboBox;
    laXMLDocumentType: TLabel;
    procedure BeforeImport(Sender: TObject);
    procedure AfterImport(Sender: TObject);
    procedure ImportRecord(Sender: TObject);
    procedure ImportError(Sender: TObject);
    procedure ImportErrorAdv(Sender: TObject);
    procedure NeedCommit(Sender: TObject);
    procedure ImportCancel(Sender: TObject; var Continue: Boolean);
    procedure BeforePost(Sender: TObject; Row: TQImportRow;
      var Accept: Boolean);
    procedure DestinationLocate(Sender: TObject; KeyColumns: TStrings;
      Row: TQImportRow; var KeyFields: string; var KeyValues: Variant);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure spbBrowseClick(Sender: TObject);
    procedure rbtClick(Sender: TObject);
    procedure edtFileNameChange(Sender: TObject);
    procedure chGoToLastPageClick(Sender: TObject);
    procedure chAutoSaveTemplateClick(Sender: TObject);
    procedure bNextClick(Sender: TObject);
    procedure bBackClick(Sender: TObject);
    procedure bDBFAddClick(Sender: TObject);
    procedure bDBFRemoveClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure bOkClick(Sender: TObject);
    procedure bHelpClick(Sender: TObject);
    procedure edtGeneratorValueChange(Sender: TObject);
    procedure edtGeneratorStepChange(Sender: TObject);
    procedure edtConstantValueChange(Sender: TObject);
    procedure edtNullValueChange(Sender: TObject);
    procedure edtDefaultValueChange(Sender: TObject);
    procedure edtLeftQuoteChange(Sender: TObject);
    procedure edtRightQuoteChange(Sender: TObject);
    procedure cmbQuoteActionChange(Sender: TObject);
    procedure cmbCharCaseChange(Sender: TObject);
    procedure cmbCharSetChange(Sender: TObject);
    procedure lstFormatFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure pgFormatsChange(Sender: TObject);
    procedure btnLoadTemplateClick(Sender: TObject);
    procedure edtShortDateFormatChange(Sender: TObject);
    procedure edtLongDateFormatChange(Sender: TObject);
    procedure edtShortTimeFormatChange(Sender: TObject);
    procedure edtLongTimeFormatChange(Sender: TObject);
    procedure chCommitAfterDoneClick(Sender: TObject);
    procedure edtCommitRecCountChange(Sender: TObject);
    procedure chImportAllRecordsClick(Sender: TObject);
    procedure edtImportRecCountChange(Sender: TObject);
    procedure chEnableErrorLogClick(Sender: TObject);
    procedure edErrorLogFileNameChange(Sender: TObject);
    procedure bErrorLogFileNameClick(Sender: TObject);
    procedure chRewriteErrorLogFileClick(Sender: TObject);
    procedure chShowErrorLogClick(Sender: TObject);
    procedure edtTXTSkipLinesChange(Sender: TObject);
    procedure btnSaveTemplateClick(Sender: TObject);
    procedure bDBFAutoFillClick(Sender: TObject);
    procedure bDBFClearClick(Sender: TObject);
    procedure btnXLSAutoFillColsClick(Sender: TObject);
    procedure btnXLSAutoFillRowsClick(Sender: TObject);
    procedure btnCSVAutoFillClick(Sender: TObject);
    procedure bCancelClick(Sender: TObject);
    procedure lstDBFDataSetChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure lstDBFChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure lstDBFMapChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure sgrCSVDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure edtCSVSkipLinesChange(Sender: TObject);
    procedure btnCSVClearClick(Sender: TObject);
    procedure sgrCSVMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure rgImportModeClick(Sender: TObject);
    procedure rgAddTypeClick(Sender: TObject);
    procedure edtDecimalSeparatorExit(Sender: TObject);
    procedure edtThousandSeparatorExit(Sender: TObject);
    procedure edtDateSeparatorExit(Sender: TObject);
    procedure edtTimeSeparatorExit(Sender: TObject);
    procedure cbQuoteExit(Sender: TObject);
    procedure cbCommaExit(Sender: TObject);
    procedure bAllToRightClick(Sender: TObject);
    procedure bOneToRirghtClick(Sender: TObject);
    procedure bOneToLeftClick(Sender: TObject);
    procedure bAllToLeftClick(Sender: TObject);
    procedure lvAvailableColumnsDblClick(Sender: TObject);
    procedure lvSelectedColumnsDblClick(Sender: TObject);
    procedure KeyColumnsDragOver(Sender, Source: TObject; X,
      Y: Integer; State: TDragState; var Accept: Boolean);
    procedure KeyColumnsDragDrop(Sender, Source: TObject; X,
      Y: Integer);
    procedure chDBFSkipDeletedClick(Sender: TObject);
    procedure chCloseAfterImportClick(Sender: TObject);
{    procedure lvXLSFieldsSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);}
    procedure tbtXLSAddRangeClick(Sender: TObject);
    procedure tbtXLSEditRangeClick(Sender: TObject);
    procedure tbtXLSMoveRangeUpClick(Sender: TObject);
    procedure tbtXLSMoveRangeDownClick(Sender: TObject);
    procedure tbtXLSDelRangeClick(Sender: TObject);
    procedure tbtXLSClearFieldRangesClick(Sender: TObject);
    procedure pbDBFAddPaint(Sender: TObject);
    procedure bDBFAddMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bDBFAddMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure pbDBFAutoFillPaint(Sender: TObject);
    procedure bDBFAutoFillMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure bDBFAutoFillMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure pbDBFRemovePaint(Sender: TObject);
    procedure bDBFRemoveMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bDBFRemoveMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure pbDBFClearPaint(Sender: TObject);
    procedure bDBFClearMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bDBFClearMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure lvCSVFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure cbCSVColNumberChange(Sender: TObject);
    procedure lvTXTFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure edXLSSkipColsChange(Sender: TObject);
    procedure edXLSSkipRowsChange(Sender: TObject);
    procedure lvXLSRangesDblClick(Sender: TObject);
    procedure bXMLAddClick(Sender: TObject);
    procedure bXMLAutoFillClick(Sender: TObject);
    procedure bXMLRemoveClick(Sender: TObject);
    procedure bXMLClearClick(Sender: TObject);
    procedure bXMLAddMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bXMLAddMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure pbXMLAddPaint(Sender: TObject);
    procedure bXMLAutoFillMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bXMLAutoFillMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure pbXMLAutoFillPaint(Sender: TObject);
    procedure bXMLRemoveMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bXMLRemoveMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure pbXMLRemovePaint(Sender: TObject);
    procedure bXMLClearMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bXMLClearMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure pbXMLClearPaint(Sender: TObject);
    procedure tbtTXTClearClick(Sender: TObject);
    procedure tbtXLSClearAllRangesClick(Sender: TObject);
    procedure lvXLSRangesChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure tbtAddReplacementClick(Sender: TObject);
    procedure lvReplacementsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure tbtEditReplacementClick(Sender: TObject);
    procedure tbtDelReplacementClick(Sender: TObject);
    procedure lvReplacementsDblClick(Sender: TObject);
    procedure lvXLSFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure chXMLWriteOnFlyClick(Sender: TObject);
    procedure lvXLSFieldsEnter(Sender: TObject);
    procedure lvXLSFieldsExit(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure cmbTXTEncodingChange(Sender: TObject);
    procedure cmbCSVEncodingChange(Sender: TObject);
    procedure cbXMLDocDataLocationChange(Sender: TObject);
    procedure chAutoTrimValueClick(Sender: TObject);
    procedure chImportEmptyRowsClick(Sender: TObject);
    procedure mScriptChange(Sender: TObject);
  private
    FImportType: TAllowedImport;
    FComma: AnsiChar;
    FQuote: AnsiChar;
    FFileName: qiString;
    FGoToLastPage: boolean;
    FAutoSaveTemplate: boolean;
    FImport: TQImport3;

    FStep: Integer;

    FDataSet: TDataSet;
    FDBGrid: TDBGrid;
    FListView: TListView;
    FStringGrid: TStringGrid;

    FImportMode: TQImportMode;
    FAddType: TQImportAddType;

    FShift: TShiftState;
    FProgress: TfmQImport3ProgressDlg;
    FTotalRecCount: Integer;
    FDataFormats: TQImportFieldFormats;
    FNeedLoadFile: boolean;
    FNeedLoadFields: boolean;

    FDBFSkipDeleted: boolean;

    FTXTItemIndex: Integer;
    FTXTSkipLines: Integer;
    FTXTEncoding: TQICharsetType;
    FTXTClearing: boolean;

    FCSVItemIndex: Integer;
    FCSVSkipLines: Integer;
    FCSVEncoding: TQICharsetType;

    FXLSFile: TXLSFile;
    FXLSSkipRows: Integer;
    FXLSSkipCols: Integer;
    FXLSIsEditingGrid: boolean;
    FXLSGridSelection: TMapRow;
    FXLSDefinedRanges: TMapRow; 

    FXMLWriteOnFly: boolean; 

    FDecimalSeparator: Char;
    FThousandSeparator: Char;
    FShortDateFormat: String;
    FLongDateFormat: String;
    FDateSeparator: Char;
    FShortTimeFormat: String;
    FLongTimeFormat: String;
    FTimeSeparator: Char;

    FCommitAfterDone: boolean;
    FCommitRecCount: Integer;
    FImportRecCount: Integer;
    FCloseAfterImport: boolean;
    FEnableErrorLog: boolean;
    FErrorLogFileName: qiString;
    FRewriteErrorLogFile: boolean;
    FShowErrorLog: boolean;
    FAutoTrimValue: Boolean;
    FImportEmptyRows: Boolean;

    FLoadingFormatItem: boolean;
    FFormatItem: TListItem;

    FTmpFileName: qiString;
    
{$IFDEF ADO}
    ImpAccess: TADO_QImport3Access;
    FAccessPassword: string;
    FAccessSourceType: TQImportAccessSourceType;
{$ENDIF}

{$IFDEF HTML}
{$IFDEF VCL6}
    impHTML: TQImport3HTML;
    FHTMLFile: THTMLFile;
    FHTMLSkipLines: Integer;
{$ENDIF}
{$ENDIF}

{$IFDEF XLSX}
{$IFDEF VCL6}
    impXlsx: TQImport3Xlsx;
    FXlsxSkipLines: Integer;
    FXlsxSheetName: qiString;
    FXlsxCurrentStringGrid: TqiStringGrid;
    FXlsxLoadHiddenSheet: Boolean;
    FXlsxNeedFillMerge: Boolean;
{$ENDIF}
{$ENDIF}

{$IFDEF DOCX}
{$IFDEF VCL6}
    impDocx: TQImport3Docx;
//    FDocxFile: TDocxFile;
    FDocxSkipLines: Integer;
    FDocxTableNumber: Integer;
    FDocxCurrentStringGrid: TqiStringGrid;
{$ENDIF}
{$ENDIF}

{$IFDEF ODS}
{$IFDEF VCL6}
    impODS: TQImport3ODS;
    FODSSkipLines: Integer;
    FODSSheetName: AnsiString;
    FODSCurrentStringGrid: TqiStringGrid;
{$ENDIF}
{$ENDIF}

{$IFDEF ODT}
{$IFDEF VCL6}
    impODT: TQImport3ODT;
    FODTSkipLines: Integer;
    FODTUseHeader: Boolean;
    FODTSheetName: AnsiString;
    FODTCurrentStringGrid: TqiStringGrid;
    FODTComplexArray: array of Boolean;
    FODTComplex: Boolean;
{$ENDIF}
{$ENDIF}

{$IFDEF XMLDOC}
{$IFDEF VCL6}
    impXMLDoc: TQImport3XMLDoc;
    FXMLDocFile: TXMLDocFile;
    FXMLDocSkipLines: Integer;
    FXMLDocXPath: qiString;
    FXMLDocDataLocation: TXMLDataLocation;
{$ENDIF}
{$ENDIF}
    {$IFDEF USESCRIPT}
    FQImport3JScriptEngine: TQImport3JScriptEngine;
    {$ENDIF}

{$IFDEF ADO}
    procedure rbtAccessTableClick(Sender: TObject);
    procedure rbtAccessSQLClick(Sender: TObject);
    procedure sbtAccessSQLLoadClick(Sender: TObject);
    procedure sbtAccessSQLSaveClick(Sender: TObject);
    procedure lbAccessTablesClick(Sender: TObject);
    procedure memAccessSQLChange(Sender: TObject);
    procedure bAccessAddClick(Sender: TObject);
    procedure bAccessAutoFillClick(Sender: TObject);
    procedure bAccessRemoveClick(Sender: TObject);
    procedure bAccessClearClick(Sender: TObject);
    procedure lstAccessMapChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure lstAccessDataSetChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure lstAccessChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure edAccessPasswordChange(Sender: TObject);
{$ENDIF}

{$IFDEF HTML}
{$IFDEF VCL6}
    procedure edtHTMLSkipLinesChange(Sender: TObject);
    procedure lvHTMLFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure sgrHTMLDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure cbHTMLColNumberChange(Sender: TObject);
    procedure cbHTMLTableNumberChange(Sender: TObject);
    procedure sgrHTMLMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure tbtHTMLAutoFillClick(Sender: TObject);
    procedure tbtHTMLClearClick(Sender: TObject);
{$ENDIF}
{$ENDIF}

{$IFDEF XLSX}
{$IFDEF VCL6}
    procedure edXlsxSkipRowsChange(Sender: TObject);
    procedure lvXlsxFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure sgrXlsxDrawCell(Sender: TObject; ACol,
      ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure sgrXlsxMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure tlbXlsxAutoFillClick(Sender: TObject);
    procedure tlbXlsxClearClick(Sender: TObject);
    procedure pcXlsxFileChange(Sender: TObject);
{$ENDIF}
{$ENDIF}

{$IFDEF DOCX}
{$IFDEF VCL6}
    procedure edDocxSkipRowsChange(Sender: TObject);
    procedure lvDocxFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure sgrDocxDrawCell(Sender: TObject; ACol,
      ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure sgrDocxMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure tlbDocxAutoFillClick(Sender: TObject);
    procedure tlbDocxClearClick(Sender: TObject);
    procedure pcDocxFileChange(Sender: TObject);
{$ENDIF}
{$ENDIF}

{$IFDEF ODS}
{$IFDEF VCL6}
    procedure edODSSkipRowsChange(Sender: TObject);
    procedure lvODSFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure sgrODSDrawCell(Sender: TObject; ACol,
      ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure sgrODSMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure tlbODSAutoFillClick(Sender: TObject);
    procedure tlbODSClearClick(Sender: TObject);
    procedure pcODSFileChange(Sender: TObject);
{$ENDIF}
{$ENDIF}

{$IFDEF ODT}
{$IFDEF VCL6}
    procedure edODTSkipRowsChange(Sender: TObject);
    procedure lvODTFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure sgrODTDrawCell(Sender: TObject; ACol,
      ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure sgrODTMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure tlbODTAutoFillClick(Sender: TObject);
    procedure tlbODTClearClick(Sender: TObject);
    procedure pcODTFileChange(Sender: TObject);
    procedure cbODTUseHeaderClick(Sender: TObject);
{$ENDIF}
{$ENDIF}

{$IFDEF XMLDOC}
{$IFDEF VCL6}
    procedure edXMLDocSkipLinesChange(Sender: TObject);
    procedure cbXMLDocColNumberChange(Sender: TObject);
    procedure sgrXMLDocDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure sgrXMLDocMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure lvXMLDocFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure tbtXMLDocAutoFillClick(Sender: TObject);
    procedure tbtXMLDocClearClick(Sender: TObject);
    procedure bXMLDocFillGridClick(Sender: TObject);
    procedure bXMLDocBuildTreeClick(Sender: TObject);
    procedure bXMLDocGetXPathClick(Sender: TObject);
{$ENDIF}
{$ENDIF}

    procedure SetImportType(const Value: TAllowedImport);
    procedure SetCaptions;
    procedure TuneVisible;
    procedure TuneOpenDialog;
    procedure ChangeExtension;
{$IFDEF ADO}
    procedure TuneAccessControls;
{$ENDIF}

    procedure TuneCSVControls;
    procedure TuneXMLControls;
    procedure SetFileName(const Value: qiString);
    procedure SetGoToLastPage(Value: boolean);
    procedure SetAutoSaveTemplate(Value: boolean);

    procedure FillCombosAndLists;
    procedure FillKeyColumns(Strings: TStrings);
    procedure MoveToSelected(Source, Destination: TListView;
      All: boolean; Index: Integer);

    function GetWizard: TQImport3Wizard;
    function GetTemplateFileName: qiString;
    function GetAutoLoadTemplate: boolean;
    function GetImportDestination: TQImportDestination;
    function GetGridCaptionRow: Integer;
    function GetGridStartRow: Integer;
    function GetKeyColumns: TStrings;

    procedure SetComma(const Value: AnsiChar);
    procedure SetQuote(const Value: AnsiChar);
    procedure SetStep(const Value: Integer);
    procedure ShowTip(Parent: TWinControl; Left, Top, Height, Width: Integer;
      const Tip: qiString);

    function AllowedImportFileType(const AFileName: qiString): Boolean;
    function ImportTypeEquivFileType(const AFileName: qiString): Boolean;
    function ImportTypeStr(AImportType: TAllowedImport): string;
    procedure CheckExtension(var Ext: qiString);

    //---- DBF page's methods
    procedure DBFFillList;
    procedure DBFClearList;
    procedure DBFFillTableList;
    procedure DBFClearTableList;
    procedure DBFTune;
    function DBFReady: boolean;

{$IFDEF HTML}
{$IFDEF VCL6}
    //---- HTML page's methods
    procedure HTMLTune;
    procedure HTMLFillList;
    procedure HTMLClearList;
    procedure HTMLFillGrid;
    procedure HTMLFillComboColumn;
    function HTMLReady: boolean;
    function HTMLCol: Integer;
{$ENDIF}
{$ENDIF}

{$IFDEF XLSX}
{$IFDEF VCL6}
    //---- Xlsx page's methods
    procedure XlsxTune;
    procedure XlsxFillList;
    procedure XlsxClearList;
    procedure XlsxFillGrid;
    procedure XlsxClearAll;
    function XlsxReady: boolean;
    function XlsxCol: Integer;
    function XlsxSheetNamingCol: qiString;
{$ENDIF}
{$ENDIF}

{$IFDEF DOCX}
{$IFDEF VCL6}
    //---- Docx page's methods
    procedure DocxTune;
    procedure DocxFillList;
    procedure DocxClearList;
    procedure DocxFillGrid;
    procedure DocxClearAll;
    function DocxReady: boolean;
    function DocxCol: Integer;
{$ENDIF}
{$ENDIF}

{$IFDEF ODS}
{$IFDEF VCL6}
    //---- ODS page's methods
    procedure ODSTune;
    procedure ODSFillList;
    procedure ODSClearList;
    procedure ODSFillGrid;
    procedure ODSClearAll;
    function ODSReady: boolean;
    function ODSCol: Integer;
{$ENDIF}
{$ENDIF}

{$IFDEF ODT}
{$IFDEF VCL6}
    //---- ODT page's methods
    procedure ODTTune;
    procedure ODTFillList;
    procedure ODTClearList;
    procedure ODTFillGrid;
    procedure ODTClearAll;
    function ODTReady: boolean;
    function ODTCol: Integer;
{$ENDIF}
{$ENDIF}

    //---- XML page's methods
    procedure XMLFillList;
    procedure XMLClearList;
    procedure XMLFillTableList;
    procedure XMLClearTableList;
    procedure XMLTune;
    function XMLReady: boolean;

{$IFDEF XMLDOC}
{$IFDEF VCL6}
    //---- XMLDoc page's methods
    procedure XMLDocTune;
    procedure XMLDocLoadFile;
    procedure XMLDocFillList;
    procedure XMLDocClearList;
    procedure XMLDocFillGrid;
    procedure XMLDocClearGrid;
    procedure XMLDocFillComboColumn;
    function XMLDocReady: boolean;
    function XMLDocCol: Integer;
    function GetXPath: qiString;
{$ENDIF}
{$ENDIF}

    //---- TXT page's methods
    procedure TXTFillCombo;
    procedure TXTClearCombo;
    procedure TXTTune;
    function TXTReady: boolean;
    procedure TXTExtractPosSize(const S: string; var Position, Size: Integer);
    procedure TXTViewerChangeSelection(Sender: TObject);
    procedure TXTViewerDeleteArrow(Sender: TObject; Position: Integer);
    procedure TXTViewerMoveArrow(Sender: TObject; OldPos, NewPos: Integer);
    procedure TXTViewerIntersectArrows(Sender: TObject; Position: Integer);

    //---- CSV page's methods
    procedure CSVFillCombo;
    procedure CSVClearCombo;
    procedure CSVTune;
    function CSVReady: boolean;
    function CSVCol: Integer;
    procedure CSVFillGrid;

    //---- XLS page's methods
    procedure XLSFillFieldList;
    procedure XLSClearFieldList;
    procedure XLSClearDataSheets;
    procedure XLSFillGrid;
    procedure XLSDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect;
      State: TGridDrawState);
    procedure XLSMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure XLSSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure XLSGridExit(Sender: TObject);
    procedure XLSGridKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure XLSGridClick(Sender: TObject);

    procedure DrawBlob(ACanvas: TCanvas; ARect: TRect);

    procedure XLSStartEditing;
    procedure XLSFinishEditing;
    procedure XLSApplyEditing;
    procedure XLSDeleteSelectedRanges;

    function XLSGetCurrentGrid: TqiStringGrid;
    procedure XLSRepaintCurrentGrid;
    procedure XLSFillSelection;

    procedure XLSTune;
    function  XLSReady: boolean;

{$IFDEF ADO}
    //---- Access page's methods
    procedure SetAccessPassword(const Value: string);
    procedure SetAccessSourceType(const Value: TQImportAccessSourceType);
    procedure AccessTune_01;
    procedure AccessTune_02;
    procedure AccessGetTableNames;
    procedure AccessFillList;
    procedure AccessGetFieldNames;
    procedure AccessClearList;
    function AccessReady: boolean;
{$ENDIF}

    //---- Formats
    procedure FormatsFillList;
    procedure FormatsClearList;
    procedure TuneFormats;
    procedure SetEnabledDataFormatControls;
    procedure ShowFormatItem(Item: TListItem);

    //---- Utilities
    procedure LoadTemplateFromFile(const AFileName: qiString);
    procedure SaveTemplateToFile(const AFileName: qiString);
    procedure SetTitle;

    procedure TuneStart;
    procedure TuneFinish;
    procedure TuneMap;
    procedure TuneButtons;
    function StartReady: boolean;

    //---- Property's methods
    procedure SetXLSSkipCols(const Value: Integer);
    procedure SetXLSSkipRows(const Value: Integer);
    procedure SetDBFSkipDeleted(const Value: boolean);
    procedure SetTXTSkipLines(const Value: Integer);
    procedure SetTXTEncoding(const Value: TQICharsetType);
    procedure SetCSVSkipLines(const Value: Integer);
    procedure SetCSVEncoding(const Value: TQICharsetType);
{$IFDEF HTML}
{$IFDEF VCL6}
    procedure SetHTMLSkipLines(const Value: Integer);
{$ENDIF}
{$ENDIF}
{$IFDEF XLSX}
{$IFDEF VCL6}
    procedure SetXlsxSkipLines(const Value: Integer);
    procedure SetXlsxSheetName(const Value: qiString);
    procedure SetXlsxLoadHiddenSheet(const Value: Boolean);
    procedure SetXlsxNeedFillMerge(const Value: Boolean);
{$ENDIF}
{$ENDIF}
{$IFDEF DOCX}
{$IFDEF VCL6}
    procedure SetDocxSkipLines(const Value: Integer);
    procedure SetDocxTableNumber(const Value: Integer);
//    procedure SetXlsxNeedFillMerge(const Value: Boolean);
{$ENDIF}
{$ENDIF}
{$IFDEF ODS}
{$IFDEF VCL6}
    procedure SetODSSkipLines(const Value: Integer);
    procedure SetODSSheetName(const Value: AnsiString);
{$ENDIF}
{$ENDIF}
{$IFDEF ODT}
{$IFDEF VCL6}
    procedure SetODTSkipLines(const Value: Integer);
    procedure SetODTSheetName(const Value: AnsiString);
{$ENDIF}
{$ENDIF}
    procedure SetXMLWriteOnFly(const Value: boolean);
{$IFDEF XMLDOC}
{$IFDEF VCL6}
    procedure SetXMLDocSkipLines(const Value: Integer);
    procedure SetXMLDocXPath(const Value: qiString);
    procedure SetXMLDocDataLocation(const Value: TXMLDataLocation);
{$ENDIF}
{$ENDIF}
    procedure SetDecimalSeparator(const Value: Char);
    procedure SetThousandSeparator(const Value: Char);
    procedure SetShortDateFormat(const Value: String);
    procedure SetLongDateFormat(const Value: String);
    procedure SetDateSeparator(const Value: Char);
    procedure SetShortTimeFormat(const Value: String);
    procedure SetLongTimeFormat(const Value: String);
    procedure SetTimeSeparator(const Value: Char);

    procedure SetCommitAfterDone(const Value: boolean);
    procedure SetCommitRecCount(const Value: Integer);
    procedure SetImportRecCount(const Value: Integer);
    procedure SetCloseAfterImport(const Value: boolean);
    procedure SetEnableErrorLog(const Value: boolean);
    procedure SetErrorLogFileName(const Value: qiString);
    procedure SetRewriteErrorLogFile(const Value: boolean);
    procedure SetShowErrorLog(const Value: boolean);

    procedure SetImportMode(const Value: TQImportMode);
    procedure SetAddType(const Value: TQImportAddType);

    procedure ApplyDataFormats(AImport: TQImport3);
    procedure SetAutoTrimValue(const Value: Boolean);
    procedure SetImportEmptyRows(const Value: Boolean);
{$IFDEF ADO}
    procedure bAccessAddMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bAccessAddMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bAccessAutoFillMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure bAccessAutoFillMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bAccessClearMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bAccessClearMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bAccessRemoveMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bAccessRemoveMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure pbAccessAddPaint(Sender: TObject);
    procedure pbAccessAutoFillPaint(Sender: TObject);
    procedure pbAccessClearPaint(Sender: TObject);
    procedure pbAccessRemovePaint(Sender: TObject);
{$ENDIF}
  protected
{$IFDEF HTML}
{$IFDEF VCL6}
    sgrHTML: TqiStringGrid;
{$ENDIF}
{$ENDIF}
{$IFDEF XMLDOC}
{$IFDEF VCL6}
    sgrXMLDoc: TqiStringGrid;
{$ENDIF}
{$ENDIF}

    sgrCSV: TqiStringGrid;
    vwTXT: TQImport3TXTViewer;
    paTip: TInfoPanel;
    property Wizard: TQImport3Wizard read GetWizard;
    property TemplateFileName: qiString read GetTemplateFileName;
    property AutoLoadTemplate: boolean read GetAutoLoadTemplate;
    property ImportDestination: TQImportDestination read GetImportDestination;
    property GridCaptionRow: Integer read GetGridCaptionRow;
    property GridStartRow: Integer read GetGridStartRow;
    property KeyColumns: TStrings read GetKeyColumns;

    property ImportType: TAllowedImport read FImportType write SetImportType;
    property FileName: qiString read FFileName write SetFileName;
    property GoToLastPage: boolean read FGoToLastPage write SetGoToLastPage;
    property AutoSaveTemplate: boolean read FAutoSaveTemplate
      write SetAutoSaveTemplate;

    property DataSet: TDataSet read FDataSet write FDataSet;
    property DBGrid: TDBGrid read FDBGrid write FDBGrid;
    property ListView: TListView read FListView write FListView;
    property StringGrid: TStringGrid read FStringGrid write FStringGrid;

    property Comma: AnsiChar read FComma write SetComma;
    property Quote: AnsiChar read FQuote write SetQuote;
    property Step: Integer read FStep write SetStep;

    property FieldFormats: TQImportFieldFormats read FDataFormats
      write FDataFormats;

    // XLS
    property XLSSkipCols: Integer read FXLSSkipCols write SetXLSSkipCols;
    property XLSSkipRows: Integer read FXLSSkipRows write SetXLSSkipRows;

{$IFDEF ADO}
    // Access
    property AccessPassword: string read FAccessPassword
      write SetAccessPassword;
    property AccessSourceType: TQImportAccessSourceType read FAccessSourceType
      write SetAccessSourceType;
{$ENDIF}

    // DBF
    property DBFSkipDeleted: boolean read FDBFSkipDeleted
      write SetDBFSkipDeleted;

    // TXT
    property TXTSkipLines: Integer read FTXTSkipLines write SetTXTSkipLines;
    property TXTEncoding: TQICharsetType read FTXTEncoding write SetTXTEncoding;

    // CSV
    property CSVSkipLines: Integer read FCSVSkipLines write SetCSVSkipLines;
    property CSVEncoding: TQICharsetType read FCSVEncoding write SetCSVEncoding;

{$IFDEF HTML}
{$IFDEF VCL6}
    // HTML
    property HTMLSkipLines: Integer read FHTMLSkipLines write SetHTMLSkipLines;
{$ENDIF}
{$ENDIF}

{$IFDEF XLSX}
{$IFDEF VCL6}
    // Xlsx
    property XlsxSkipLines: Integer read FXlsxSkipLines write SetXlsxSkipLines;
    property XlsxSheetName: qiString read FXlsxSheetName write SetXlsxSheetName;
    property XlsxNeedFillMerge: Boolean read FXlsxNeedFillMerge write SetXlsxNeedFillMerge;
    property XlsxLoadHiddenSheet: Boolean read FXlsxLoadHiddenSheet write SetXlsxLoadHiddenSheet;
{$ENDIF}
{$ENDIF}

{$IFDEF DOCX}
{$IFDEF VCL6}
    // Docx
    property DocxSkipLines: Integer read FDocxSkipLines write SetDocxSkipLines;
    property DocxTableNumber: Integer read FDocxTableNumber write SetDocxTableNumber;
{$ENDIF}
{$ENDIF}

{$IFDEF ODS}
{$IFDEF VCL6}
    // ODS
    property ODSSkipLines: Integer read FODSSkipLines write SetODSSkipLines;
    property ODSSheetName: AnsiString read FODSSheetName write SetODSSheetName;
{$ENDIF}
{$ENDIF}

{$IFDEF ODT}
{$IFDEF VCL6}
    // ODT
    property ODTSkipLines: Integer read FODTSkipLines write SetODTSkipLines;
    property ODTSheetName: AnsiString read FODTSheetName write SetODTSheetName;
    property ODTUseHeader: Boolean read FODTUseHeader write FODTUseHeader;
    property ODTComplex: Boolean read FODTComplex write FODTComplex;
{$ENDIF}
{$ENDIF}

    // XML
    property XMLWriteOnFly: boolean read FXMLWriteOnFly write SetXMLWriteOnFly;

{$IFDEF XMLDOC}
{$IFDEF VCL6}
    // XMLDoc
    property XMLDocSkipLines: Integer read FXMLDocSkipLines
      write SetXMLDocSkipLines;
    property XMLDocXPath: qiString read FXMLDocXPath write SetXMLDocXPath;
    property XMLDocDataLocation: TXMLDataLocation read FXMLDocDataLocation
      write SetXMLDocDataLocation;
{$ENDIF}
{$ENDIF}

    // Base format
    property DecimalSeparator: Char read FDecimalSeparator
      write SetDecimalSeparator;
    property ThousandSeparator: Char read FThousandSeparator
      write SetThousandSeparator;
    property ShortDateFormat: String read FShortDateFormat
      write SetShortDateFormat;
    property LongDateFormat: String read FLongDateFormat
      write SetLongDateFormat;
    property DateSeparator: Char read FDateSeparator write SetDateSeparator;
    property ShortTimeFormat: String read FShortTimeFormat
      write SetShortTimeFormat;
    property LongTimeFormat: String read FLongTimeFormat
      write SetLongTimeFormat;
    property TimeSeparator: Char read FTimeSeparator write SetTimeSeparator;

    //---- Last Step
    property CommitAfterDone: boolean read FCommitAfterDone
      write SetCommitAfterDone;
    property CommitRecCount: Integer read FCommitRecCount
      write SetCommitRecCount;
    property ImportRecCount: Integer read FImportRecCount
      write SetImportRecCount;
    property CloseAfterImport: boolean read FCloseAfterImport
      write SetCloseAfterImport;
    property EnableErrorLog: boolean read FEnableErrorLog
      write SetEnableErrorLog;
    property ErrorLogFileName: qiString read FErrorLogFileName
      write SetErrorLogFileName;
    property RewriteErrorLogFile: boolean read FRewriteErrorLogFile
      write SetRewriteErrorLogFile;
    property ShowErrorLog: boolean read FShowErrorLog
      write SetShowErrorLog;
    property AutoTrimValue: Boolean read FAutoTrimValue
      write SetAutoTrimValue;
    property ImportEmptyRows: Boolean read FImportEmptyRows
      write SetImportEmptyRows;

    property ImportMode: TQImportMode read FImportMode
      write SetImportMode;
    property AddType: TQImportAddType read FAddType
      write SetAddType;
  public
    property Import: TQImport3 read FImport write FImport;
  end;

implementation

uses
  {$IFDEF VCL16}
    System.Variants,
    System.Math,
    System.SysUtils,
    Winapi.Messages,
    System.IniFiles,
    {$IFDEF VCL17}
      System.UITypes,
    {$ENDIF}
  {$ELSE}
    {$IFDEF VCL6}
      Variants,
    {$ENDIF}
    Math,
    SysUtils,
    Messages,
    IniFiles,
  {$ENDIF}
  {$IFDEF QI_UNICODE}
    QImport3GpTextFile,
  {$ENDIF}
  QImport3StrIDs,
  QImport3DBFFile,
  fuQImport3Loading,
  QImport3Common,
  QImport3XLSUtils,
  QImport3XLSCalculate,
  fuQImport3XLSRangeEdit,
  fuQImport3ReplacementEdit,
  QImport3WideStringCanvas,
  QImport3XlsxMapParser;

{$R *.DFM}

{ TQImport3Wizard }

constructor TQImport3Wizard.Create(AOwner: TComponent);
begin
  inherited;
  FAllowedImports := [Low(TAllowedImport)..High(TAllowedImport)];
  FImportRecCount := 0;
  FCommitRecCount := 1000;
  FCommitAfterDone := False;
  FErrorLog := False;
  FErrorLogFileName := 'error.log';
  FRewriteErrorLogFile := True;
  FShowErrorLog := False;
  FShowProgress := True;
  FAutoChangeExtension := True;
  FShowHelpButton := True;
  FCloseAfterImport := False;
  FFormats := TQImportFormats.Create;
  FFieldFormats := TQImportFieldFormats.Create(Self);

  FShowSaveLoadButtons := True;
  FAutoLoadTemplate := False;
  FAutoSaveTemplate := False;
  FGoToLastPage := False;

  FImportDestination := qidDataSet;
  FImportMode := qimInsertAll;
  FAddType := qatInsert;
  FKeyColumns := TStringList.Create;
  FGridCaptionRow := -1;
  FGridStartRow := -1;
  FConfirmOnCancel := True;

  FPicture := TPicture.Create;

  FTextViewerRows := 20;
  FCSVViewerRows := 20;
  FExcelViewerRows := 256;
  FExcelMaxColWidth := 130;
  FAutoTrimValue := False;
  FImportEmptyRows := True;
end;

destructor TQImport3Wizard.Destroy;
begin
  FPicture.Free;
  FFieldFormats.Free;
  FFormats.Free;
  FKeyColumns.Free;
  inherited;
end;

procedure TQImport3Wizard.Execute;
begin
{$IFNDEF HTML}
{$IFNDEF VCL6}
  Exclude(FAllowedImports, aiHTML);
{$ENDIF}{$ENDIF}

{$IFNDEF XLSX}
{$IFNDEF VCL6}
  Exclude(FAllowedImports, aiXlsx);
{$ENDIF}{$ENDIF}

{$IFNDEF DOCX}
{$IFNDEF VCL6}
  Exclude(FAllowedImports, aiDocx);
{$ENDIF}{$ENDIF}

{$IFNDEF ODS}
{$IFNDEF VCL6}
  Exclude(FAllowedImports, aiODS);
{$ENDIF}{$ENDIF}

{$IFNDEF ODT}
{$IFNDEF VCL6}
  Exclude(FAllowedImports, aiODT);
{$ENDIF}{$ENDIF}

{$IFNDEF XMLDOC}
{$IFNDEF VCL6}
  Exclude(FAllowedImports, aiXMLDoc);
{$ENDIF}{$ENDIF}

{$IFNDEF ADO}
  Exclude(FAllowedImports, aiAccess);
{$ENDIF}

  if AllowedImports = [] then
    raise EQImportError.Create(QImportLoadStr(QIE_AllowedImportsEmpty));

  if ImportDestination = qidUserDefined then
    raise EQImportError.Create(QImportLoadStr(QIW_DontWorkUserDefined));

  QImportCheckDestination(False, ImportDestination, DataSet, DBGrid, ListView,
    StringGrid);
  with TQImport3WizardF.Create(Self) do
    try
      ShowModal;
    finally
      Free;
    end;
end;

procedure TQImport3Wizard.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited;
  if (Operation = opRemove) and (AComponent = FDataSet)
    then FDataSet := nil;
  if (Operation = opRemove) and (AComponent = FDBGrid)
    then FDBGrid := nil;
  if (Operation = opRemove) and (AComponent = FListView)
    then FListView := nil;
  if (Operation = opRemove) and (AComponent = FStringGrid)
    then FStringGrid := nil;
end;

function TQImport3Wizard.IsFileName: Boolean;
begin
  Result := FFileName <> EmptyStr;
end;

procedure TQImport3Wizard.SetFormats(const Value: TQImportFormats);
begin
  FFormats.Assign(Value);
end;

procedure TQImport3Wizard.SetFieldFormats(const Value: TQImportFieldFormats);
begin
  FFieldFormats.Assign(Value);
end;

procedure TQImport3Wizard.SetKeyColumns(const Value: TStrings);
begin
  FKeyColumns.Assign(Value);
end;

procedure TQImport3Wizard.SetPicture(const Value: TPicture);
begin
  FPicture.Assign(Value);
end;

{ TQImportWizardF }

const
  FileExts: array[0..11] of string = ('.xls', '.dbf', '.xml', '.txt', '.csv', '.mdb', '.html', '.xml', '.xlsx', '.docx', '.ods', '.odt');

procedure TQImport3WizardF.BeforeImport(Sender: TObject);
begin
  FTotalRecCount := (Sender as TQImport3).TotalRecCount;
  if Assigned(FProgress) then begin
    PostMessage(FProgress.Handle, WM_QIMPORT_PROGRESS, QIP_STATE, 1);
    PostMessage(FProgress.Handle, WM_QIMPORT_PROGRESS, QIP_ROWCOUNT, FTotalRecCount);
    Application.ProcessMessages;
  end;
  if Assigned(Wizard.OnBeforeImport) then Wizard.OnBeforeImport(Wizard);
end;

procedure TQImport3WizardF.AfterImport(Sender: TObject);
begin
  if Assigned(FProgress) then begin
    PostMessage(FProgress.Handle, WM_QIMPORT_PROGRESS, QIP_FINISH, Integer(ShowErrorLog));
    if not Import.Canceled then
      PostMessage(FProgress.Handle, WM_QIMPORT_PROGRESS, QIP_STATE, 3);
    Application.ProcessMessages;
  end;
  if Assigned(Wizard.OnAfterImport) then Wizard.OnAfterImport(Wizard);
end;

procedure TQImport3WizardF.ImportRecord(Sender: TObject);
begin
  if Assigned(FProgress) then begin
    PostMessage(FProgress.Handle, WM_QIMPORT_PROGRESS, QIP_IMPORT,
      Integer(Import.LastAction));
    Application.ProcessMessages;
  end;
  if Assigned(Wizard.OnImportRecord) then
    Wizard.OnImportRecord(Wizard);
end;

procedure TQImport3WizardF.ImportError(Sender: TObject);
begin
  if Assigned(FProgress) then
  begin
    PostMessage(FProgress.Handle, WM_QIMPORT_PROGRESS, QIP_ERROR, 1);
    Application.ProcessMessages;
  end;
  if Assigned(Wizard.OnImportRecord) then
    Wizard.OnImportRecord(Wizard);
  if Assigned(Wizard.OnImportError) then
    Wizard.OnImportError(Wizard);
end;

procedure TQImport3WizardF.ImportErrorAdv(Sender: TObject);
begin
  if Assigned(FProgress) then
  begin
    PostMessage(FProgress.Handle, WM_QIMPORT_PROGRESS, QIP_ERROR_ADV, 1);
    Application.ProcessMessages;
  end;
  if Assigned(Wizard.OnImportErrorAdv) then
    Wizard.OnImportErrorAdv(Wizard);
end;

procedure TQImport3WizardF.NeedCommit(Sender: TObject);
begin
  if Assigned(FProgress) then begin
    PostMessage(FProgress.Handle, WM_QIMPORT_PROGRESS, QIP_COMMIT, 0);
    Application.ProcessMessages;
  end;
  if Assigned(Wizard.OnNeedCommit) then Wizard.OnNeedCommit(Self);
end;

procedure TQImport3WizardF.ImportCancel(Sender: TObject; var Continue: Boolean);
begin
  if Assigned(FProgress) then begin
    PostMessage(FProgress.Handle, WM_QIMPORT_PROGRESS, QIP_STATE, 4);
    Application.ProcessMessages;
  end;
  Continue := Application.MessageBox(PChar(QImportLoadStr(QIW_NeedCancel)),
                            PChar(QImportLoadStr(QIW_NeedCancelCaption)),
                            MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) = ID_NO;
  if Assigned(Wizard.OnImportCancel) then Wizard.OnImportCancel(Wizard, Continue);
  if Assigned(FProgress) then begin
    if Continue then
      PostMessage(FProgress.Handle, WM_QIMPORT_PROGRESS, QIP_STATE, 5)
    else PostMessage(FProgress.Handle, WM_QIMPORT_PROGRESS, QIP_STATE, 2);
    Application.ProcessMessages;
  end;
end;

procedure TQImport3WizardF.BeforePost(Sender: TObject;
  Row: TQImportRow; var Accept: Boolean);
begin
  if Assigned(Wizard.OnBeforePost) then
    Wizard.OnBeforePost(Sender, Row, Accept);
end;

procedure TQImport3WizardF.DestinationLocate(Sender: TObject;
  KeyColumns: TStrings; Row: TQImportRow; var KeyFields: string;
  var KeyValues: Variant);
begin
  if Assigned(Wizard.OnDestinationLocate) then
    Wizard.OnDestinationLocate(Sender, KeyColumns, Row, KeyFields, KeyValues);
end;

procedure TQImport3WizardF.DrawBlob(ACanvas: TCanvas; ARect: TRect);
var
  Bmp: TBitmap;
begin
  Bmp := TBitmap.Create;
  try
    ilWizard.GetBitmap(22, Bmp);
    ACanvas.Draw(ARect.Left +
      Round((ARect.Right - ARect.Left)/2 - Bmp.Width/2), ARect.Top, Bmp);
  finally
    Bmp.Free;
  end;
end;

procedure TQImport3WizardF.FormCreate(Sender: TObject);
var
  i: Integer;
begin
  pgImport.ActivePage := tsImportType;
  for i := 0 to pgImport.PageCount - 1 do
    pgImport.Pages[i].Parent := Self;

  TuneVisible;
  {$IFDEF ADO}
  ImpAccess := TADO_QImport3Access.Create(Self);
  ImpAccess.AddType := qatInsert;
  ImpAccess.OnBeforeImport := BeforeImport;
  ImpAccess.OnAfterImport := AfterImport;
  ImpAccess.OnImportRecord := ImportRecord;
  ImpAccess.OnImportError := ImportError;
  ImpAccess.OnImportErrorAdv := ImportErrorAdv;
  ImpAccess.OnNeedCommit := NeedCommit;
  ImpAccess.OnImportCancel := ImportCancel;
  ImpAccess.OnBeforePost := BeforePost;
  ImpAccess.OnDestinationLocate := DestinationLocate;

  rbtAccessTable.OnClick := rbtAccessTableClick;
  lbAccessTables.OnClick := lbAccessTablesClick;
  rbtAccessSQL.OnClick := rbtAccessSQLClick;
  tbtAccessSQLLoad.OnClick := sbtAccessSQLLoadClick;
  tbtAccessSQLSave.OnClick := sbtAccessSQLSaveClick;
  memAccessSQL.OnChange := memAccessSQLChange;
  lstAccessDataSet.OnChange := lstAccessDataSetChange;
  lstAccess.OnChange := lstAccessChange;
  lstAccessMap.OnChange := lstAccessMapChange;
  bAccessAdd.OnClick := bAccessAddClick;
  bAccessAdd.OnMouseDown := bAccessAddMouseDown;
  bAccessAdd.OnMouseUp := bAccessAddMouseUp;
  bAccessAutoFill.OnClick := bAccessAutoFillClick;
  bAccessAutoFill.OnMouseDown := bAccessAutoFillMouseDown;
  bAccessAutoFill.OnMouseUp := bAccessAutoFillMouseUp;
  bAccessRemove.OnClick := bAccessRemoveClick;
  bAccessRemove.OnMouseDown := bAccessRemoveMouseDown;
  bAccessRemove.OnMouseUp := bAccessRemoveMouseUp;
  bAccessClear.OnClick := bAccessClearClick;
  bAccessClear.OnMouseDown := bAccessClearMouseDown;
  bAccessClear.OnMouseUp := bAccessClearMouseUp;
  pbAccessAdd.OnPaint := pbAccessAddPaint;
  pbAccessAutoFill.OnPaint := pbAccessAutoFillPaint;
  pbAccessRemove.OnPaint := pbAccessRemovePaint;
  pbAccessClear.OnPaint := pbAccessClearPaint;
  edAccessPassword.OnChange := edAccessPasswordChange;
  {$ENDIF}

  {$IFDEF HTML}
  {$IFDEF VCL6}
  impHTML := TQImport3HTML.Create(Self);
  impHTML.AddType := qatInsert;
  impHTML.OnBeforeImport := BeforeImport;
  impHTML.OnAfterImport := AfterImport;
  impHTML.OnImportRecord := ImportRecord;
  impHTML.OnImportError := ImportError;
  impHTML.OnImportErrorAdv := ImportErrorAdv;
  impHTML.OnNeedCommit := NeedCommit;
  impHTML.OnImportCancel := ImportCancel;
  impHTML.OnBeforePost := BeforePost;
  impHTML.OnDestinationLocate := DestinationLocate;

  sgrHTML := TqiStringGrid.Create(Self);
  sgrHTML.Parent := tsHTMLOptions;
  sgrHTML.Left := tbHTML.Left;
  sgrHTML.Top := tbHTML.Top + tbHTML.Height + 8;
  sgrHTML.Width := Width - (Constraints.MinWidth - 402);
  sgrHTML.Height := Height - (Constraints.MinHeight - 280); 
  sgrHTML.ColCount := 1;
  sgrHTML.DefaultRowHeight := 16;
  sgrHTML.FixedCols := 0;
  sgrHTML.RowCount := 1;
  sgrHTML.FixedRows := 0;
  sgrHTML.Font.Charset := DEFAULT_CHARSET;
  sgrHTML.Font.Color := clWindowText;
  sgrHTML.Font.Height := -11;
  sgrHTML.Font.Name := 'Courier New';
  sgrHTML.Font.Style := [];
  sgrHTML.Options := [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine];
  sgrHTML.ParentFont := False;
  sgrHTML.TabOrder := 5;
  sgrHTML.OnDrawCell := sgrHTMLDrawCell;
  sgrHTML.OnMouseDown := sgrHTMLMouseDown;
  sgrHTML.Anchors := [akTop, akLeft, akRight, akBottom];
  lvHTMLFields.OnChange := lvHTMLFieldsChange;
  tbtHTMLAutoFill.OnClick := tbtHTMLAutoFillClick;
  tbtHTMLClear.OnClick := tbtHTMLClearClick;
  cbHTMLTableNumber.OnChange := cbHTMLTableNumberChange;
  cbHTMLColNumber.OnChange := cbHTMLColNumberChange;
  edtHTMLSkipLines.OnChange := edtHTMLSkipLinesChange;

  FHTMLFile := THTMLFile.Create;
  HTMLSkipLines := 0;
  {$ENDIF}
  {$ENDIF}

  {$IFDEF XLSX}
  {$IFDEF VCL6}
  impXlsx := TQImport3Xlsx.Create(Self);
  impXlsx.AddType := qatInsert;
  impXlsx.OnBeforeImport := BeforeImport;
  impXlsx.OnAfterImport := AfterImport;
  impXlsx.OnImportRecord := ImportRecord;
  impXlsx.OnImportError := ImportError;
  impXlsx.OnImportErrorAdv := ImportErrorAdv;
  impXlsx.OnNeedCommit := NeedCommit;
  impXlsx.OnImportCancel := ImportCancel;
  impXlsx.OnBeforePost := BeforePost;
  impXlsx.OnDestinationLocate := DestinationLocate;

  lvXlsxFields.OnChange := lvXlsxFieldsChange;
{  sgrHTML.OnDrawCell := sgrHTMLDrawCell;
  sgrHTML.OnMouseDown := sgrHTMLMouseDown;}
  tlbXlsxAutoFill.OnClick := tlbXlsxAutoFillClick;
  tlbXlsxClear.OnClick := tlbXlsxClearClick;
  edXlsxSkipRows.OnChange := edXlsxSkipRowsChange;
  pcXlsxFile.OnChange := pcXlsxFileChange;

  XlsxSkipLines := 0;
  XlsxSheetName := '';
  {$ENDIF}
  {$ENDIF}

  {$IFDEF DOCX}
  {$IFDEF VCL6}
  impDocx := TQImport3Docx.Create(Self);
  impDocx.AddType := qatInsert;
  impDocx.OnBeforeImport := BeforeImport;
  impDocx.OnAfterImport := AfterImport;
  impDocx.OnImportRecord := ImportRecord;
  impDocx.OnImportError := ImportError;
  impDocx.OnImportErrorAdv := ImportErrorAdv;
  impDocx.OnNeedCommit := NeedCommit;
  impDocx.OnImportCancel := ImportCancel;
  impDocx.OnBeforePost := BeforePost;
  impDocx.OnDestinationLocate := DestinationLocate;

  lvDocxFields.OnChange := lvDocxFieldsChange;
{  sgrDocx.OnDrawCell := sgrDocxDrawCell;
  sgrDocx.OnMouseDown := sgrDocxMouseDown;}
  tlbDocxAutoFill.OnClick := tlbDocxAutoFillClick;
  tlbDocxClear.OnClick := tlbDocxClearClick;
  edDocxSkipRows.OnChange := edDocxSkipRowsChange;
  pcDocxFile.OnChange := pcDocxFileChange;

  DocxSkipLines := 0;
  DocxTableNumber := 0;
  {$ENDIF}
  {$ENDIF}

  {$IFDEF ODS}
  {$IFDEF VCL6}
  impODS := TQImport3ODS.Create(Self);
  impODS.AddType := qatInsert;
  impODS.OnBeforeImport := BeforeImport;
  impODS.OnAfterImport := AfterImport;
  impODS.OnImportRecord := ImportRecord;
  impODS.OnImportError := ImportError;
  impODS.OnImportErrorAdv := ImportErrorAdv;
  impODS.OnNeedCommit := NeedCommit;
  impODS.OnImportCancel := ImportCancel;
  impODS.OnBeforePost := BeforePost;
  impODS.OnDestinationLocate := DestinationLocate;

  lvODSFields.OnChange := lvODSFieldsChange;
  tlbODSAutoFill.OnClick := tlbODSAutoFillClick;
  tlbODSClear.OnClick := tlbODSClearClick;
  edODSSkipRows.OnChange := edODSSkipRowsChange;
  pcODSFile.OnChange := pcODSFileChange;

  ODSSkipLines := 0;
  ODSSheetName := '';
  {$ENDIF}
  {$ENDIF}

  {$IFDEF ODT}
  {$IFDEF VCL6}
  impODT := TQImport3ODT.Create(Self);
  impODT.AddType := qatInsert;
  impODT.OnBeforeImport := BeforeImport;
  impODT.OnAfterImport := AfterImport;
  impODT.OnImportRecord := ImportRecord;
  impODT.OnImportError := ImportError;
  impODT.OnImportErrorAdv := ImportErrorAdv;
  impODT.OnNeedCommit := NeedCommit;
  impODT.OnImportCancel := ImportCancel;
  impODT.OnBeforePost := BeforePost;
  impODT.OnDestinationLocate := DestinationLocate;

  lvODTFields.OnChange := lvODTFieldsChange;
  tlbODTAutoFill.OnClick := tlbODTAutoFillClick;
  tlbODTClear.OnClick := tlbODTClearClick;
  edODTSkipRows.OnChange := edODTSkipRowsChange;
  pcODTFile.OnChange := pcODTFileChange;
  cbODTUseHeader.OnClick := cbODTUseHeaderClick;

  ODTSkipLines := 0;
  ODTSheetName := '';
  ODTUseHeader := false;
  {$ENDIF}
  {$ENDIF}
  
  {$IFDEF XMLDOC}
  {$IFDEF VCL6}
  impXMLDoc := TQImport3XMLDoc.Create(Self);
  impXMLDoc.AddType := qatInsert;
  impXMLDoc.OnBeforeImport := BeforeImport;
  impXMLDoc.OnAfterImport := AfterImport;
  impXMLDoc.OnImportRecord := ImportRecord;
  impXMLDoc.OnImportError := ImportError;
  impXMLDoc.OnImportErrorAdv := ImportErrorAdv;
  impXMLDoc.OnNeedCommit := NeedCommit;
  impXMLDoc.OnImportCancel := ImportCancel;
  impXMLDoc.OnBeforePost := BeforePost;
  impXMLDoc.OnDestinationLocate := DestinationLocate;

  sgrXMLDoc := TqiStringGrid.Create(Self);
  sgrXMLDoc.Parent := tsXMLDocOptions;
  sgrXMLDoc.Left := tbXMLDoc.Left;
  sgrXMLDoc.Top := tbXMLDoc.Top + tbXMLDoc.Height + 8;
  sgrXMLDoc.Width := Width - (Constraints.MinWidth - 360);
  sgrXMLDoc.Height := Height - (Constraints.MinHeight - 234);
  sgrXMLDoc.ColCount := 2;
  sgrXMLDoc.DefaultColWidth := 82;
  sgrXMLDoc.DefaultRowHeight := 16;
  sgrXMLDoc.RowCount := 2;
  sgrXMLDoc.Font.Charset := DEFAULT_CHARSET;
  sgrXMLDoc.Font.Color := clWindowText;
  sgrXMLDoc.Font.Height := -11;
  sgrXMLDoc.Font.Name := 'Courier New';
  sgrXMLDoc.Font.Style := [];
  sgrXMLDoc.Options := [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine];
  sgrXMLDoc.ParentFont := False;
  sgrXMLDoc.TabOrder := 1;
  sgrXMLDoc.OnDrawCell := sgrXMLDocDrawCell;
  sgrXMLDoc.OnMouseDown := sgrXMLDocMouseDown;
  sgrXMLDoc.Anchors := [akTop, akLeft, akRight, akBottom];

  lvXMLDocFields.OnChange := lvXMLDocFieldsChange;
  bXMLDocBuildTree.OnClick := bXMLDocBuildTreeClick;
  bXMLDocGetXPath.OnClick := bXMLDocGetXPathClick;
  miGetXPath.OnClick := bXMLDocGetXPathClick;
  tvXMLDoc.OnDblClick := bXMLDocGetXPathClick;
  bXMLDocFillGrid.OnClick := bXMLDocFillGridClick;
  tbtXMLDocAutoFill.OnClick := tbtXMLDocAutoFillClick;
  tbtXMLDocClear.OnClick := tbtXMLDocClearClick;
  cbXMLDocColNumber.OnChange := cbXMLDocColNumberChange;
  edXMLDocSkipLines.OnChange := edXMLDocSkipLinesChange;
  cbXMLDocDataLocation.ItemIndex := 0;

  FXMLDocFile := TXMLDocFile.Create;
  XMLDocSkipLines := 0;
  {$ENDIF}
  {$ENDIF}

  sgrCSV := TqiStringGrid.Create(Self);
  sgrCSV.Parent := tsCSVOptions;
  sgrCSV.Left := tbCSV.Left;
  sgrCSV.Top := cmbCSVEncoding.Top + cmbCSVEncoding.Height + 8;
  sgrCSV.Width := Width - sgrCSV.Left - 8;
  sgrCSV.Height := lvCSVFields.Height - sgrCSV.Top + lvCSVFields.Top + 8;
  sgrCSV.DefaultRowHeight := 16;
  sgrCSV.FixedCols := 0;
  sgrCSV.RowCount := 21;
  sgrCSV.Font.Charset := DEFAULT_CHARSET;
  sgrCSV.Font.Color := clWindowText;
  sgrCSV.Font.Height := -11;
  sgrCSV.Font.Name := 'Courier New';
  sgrCSV.Font.Style := [];
  sgrCSV.Options := [goFixedVertLine, goFixedHorzLine, goVertLine, goColSizing];
  sgrCSV.ParentFont := False;
  sgrCSV.TabOrder := 0;
  sgrCSV.OnDrawCell := sgrCSVDrawCell;
  sgrCSV.OnMouseDown := sgrCSVMouseDown;
  sgrCSV.Anchors := [akTop, akLeft, akRight, akBottom];

  vwTXT := TQImport3TXTViewer.Create(Self);
  vwTXT.Import := impASCII;
  vwTXT.Parent := pgImport.Pages[1];
  vwTXT.Height := Height - (Constraints.MinHeight - 286);
  vwTXT.Left := tbTXT.Left;
  vwTXT.Top := tbTXT.Top + tbTXT.Height + 8;
  vwTXT.Width := Width - (Constraints.MinWidth - 370);
  vwTXT.CodePage := SystemCTNames[ctWinDefined];
  vwTXT.OnChangeSelection := TXTViewerChangeSelection;
  vwTXT.OnDeleteArrow := TXTViewerDeleteArrow;
  vwTXT.OnMoveArrow := TXTViewerMoveArrow;
  vwTXT.OnIntersectArrows := TXTViewerIntersectArrows;
  vwTXT.Anchors := [akTop, akLeft, akRight, akBottom];


  paTip := TInfoPanel.Create(Self);
  paTip.AutoSize := False;
  paTip.WordWrap := True;
  paTip.Anchors := [akLeft, akTop, akRight];

  bHelp.Enabled := Wizard.ShowHelpButton;
  bHelp.Visible := Wizard.ShowHelpButton;
  btnLoadTemplate.Enabled := Wizard.ShowSaveLoadButtons;
  btnLoadTemplate.Visible := Wizard.ShowSaveLoadButtons;
  btnSaveTemplate.Enabled := Wizard.ShowSaveLoadButtons;
  btnSaveTemplate.Visible := Wizard.ShowSaveLoadButtons;

  laTemplateFileNameTag.Visible :=
    not Wizard.ShowSaveLoadButtons and Wizard.AutoLoadTemplate;

  FileName := Wizard.FileName;
  GoToLastPage := Wizard.GoToLastPage;
  AutoSaveTemplate := Wizard.AutoSaveTemplate;
  FieldFormats := Wizard.FieldFormats;

  DataSet := Wizard.DataSet;
  DBGrid := Wizard.DBGrid;
  ListView := Wizard.ListView;
  StringGrid := Wizard.StringGrid;
  FillCombosAndLists;

  TuneOpenDialog;
  ChangeExtension;
{$IFDEF ADO}
  TuneAccessControls;
{$ENDIF}
  TuneCSVControls;
  TuneXMLControls;

  rbtXLS.Enabled := aiXLS in Wizard.AllowedImports;
  rbtXlsx.Enabled := aiXlsx in Wizard.AllowedImports;
  rbtDocx.Enabled := aiDocx in Wizard.AllowedImports;
  rbtODS.Enabled := aiODS in Wizard.AllowedImports;
  rbtODT.Enabled := aiODT in Wizard.AllowedImports;
  rbtDBF.Enabled := aiDBF in Wizard.AllowedImports;
  rbtHTML.Enabled := aiHTML in Wizard.AllowedImports;
  rbtXML.Enabled := aiXML in Wizard.AllowedImports;
  rbtXMLDoc.Enabled := aiXMLDoc in Wizard.AllowedImports;
  rbtTXT.Enabled := aiTXT in Wizard.AllowedImports;
  rbtCSV.Enabled := aiCSV in Wizard.AllowedImports;
  rbtAccess.Enabled := aiAccess in Wizard.AllowedImports;

  if rbtXLS.Enabled then rbtXLS.Checked := true
  else if rbtXlsx.Enabled then rbtXlsx.Checked := true
  else if rbtDocx.Enabled then rbtDocx.Checked := true
  else if rbtODS.Enabled then rbtODS.Checked := true
  else if rbtODT.Enabled then rbtODT.Checked := true
  else if rbtDBF.Enabled then rbtDBF.Checked := true
  else if rbtHTML.Enabled then rbtHTML.Checked := true
  else if rbtXML.Enabled then rbtXML.Checked := true
  else if rbtXMLDoc.Enabled then rbtXMLDoc.Checked := true
  else if rbtTXT.Enabled then rbtTXT.Checked := true
  else if rbtCSV.Enabled then rbtCSV.Checked := true
  else if rbtAccess.Enabled then rbtAccess.Checked := true;

  FXLSFile := TxlsFile.Create;
  XLSSkipCols := 0;
  XLSSkipRows := 0;

  DBFSkipDeleted := True;

  TXTSkipLines := 0;
  FTXTEncoding := ctWinDefined;
  FTXTClearing := False;
  CSVSkipLines := 0;
  FCSVEncoding:=ctWinDefined;
  XMLWriteOnFly := False;

  DecimalSeparator := Wizard.Formats.DecimalSeparator;
  ThousandSeparator := Wizard.Formats.ThousandSeparator;
  ShortDateFormat := Wizard.Formats.ShortDateFormat;
  LongDateFormat := Wizard.Formats.LongDateFormat;
  DateSeparator := Wizard.Formats.DateSeparator;
  ShortTimeFormat := Wizard.Formats.ShortTimeFormat;
  LongTimeFormat := Wizard.Formats.LongTimeFormat;
  TimeSeparator := Wizard.Formats.TimeSeparator;

  CommitAfterDone := Wizard.CommitAfterDone;
  CommitRecCount := Wizard.CommitRecCount;
  ImportRecCount := Wizard.ImportRecCount;
  CloseAfterImport := Wizard.CloseAfterImport;
  EnableErrorLog := Wizard.ErrorLog;
  ErrorLogFileName := Wizard.ErrorLogFileName;
  RewriteErrorLogFile := Wizard.RewriteErrorLogFile;
  ShowErrorLog := Wizard.ShowErrorLog;
  chEnableErrorLogClick(nil);
  AutoTrimValue := Wizard.AutoTrimValue;
  ImportEmptyRows := Wizard.ImportEmptyRows;

  AddType := Wizard.AddType;
  ImportMode := Wizard.ImportMode;

  Step := 0;
  FTXTItemIndex := -1;
  FCSVItemIndex := -1;
  FNeedLoadFile := True;
  FNeedLoadFields := true;

  FFormatItem := nil;
  FLoadingFormatItem := False;

  if AutoLoadTemplate then
  begin
    if Assigned(Wizard.OnLoadTemplate) then
      Wizard.OnLoadTemplate(Wizard, TemplateFileName);
    LoadTemplateFromFile(TemplateFileName);
  end;

  if Assigned(Wizard.Picture.Graphic) and not Wizard.Picture.Graphic.Empty then
    Image.Picture := Wizard.Picture;

  FXLSIsEditingGrid := False;
  FXLSGridSelection := TMapRow.Create(nil);
  FXLSDefinedRanges := TMapRow.Create(nil);

  {$IFDEF USESCRIPT}
    FQImport3JScriptEngine := TQImport3JScriptEngine.Create(Self);
  {$ELSE}
    mScript.Visible := False;
    lScript.Visible := False;
    laReplacements.Top := lScript.Top;
    lvReplacements.Top := mScript.Top;
    tbReplacements.Top := mScript.Top + 20;
  {$ENDIF}
end;

procedure TQImport3WizardF.FormDestroy(Sender: TObject);
var
  i: Integer;
begin
  FXLSDefinedRanges.Free; 
  FXLSGridSelection.Free;
  FXLSFile.Free;

  {$IFDEF HTML}
  {$IFDEF VCL6}
  HTMLClearList;
  FHTMLFile.Free;
  impHTML.Free;
  {$ENDIF}
  {$ENDIF}
  
  {$IFDEF XLSX}
  {$IFDEF VCL6}
  XlsxClearAll;
  impXlsx.Free;
  {$ENDIF}
  {$ENDIF}

  {$IFDEF DOCX}
  {$IFDEF VCL6}
  DocxClearAll;
  impDocx.Free;
  {$ENDIF}
  {$ENDIF}

  {$IFDEF ODS}
  {$IFDEF VCL6}
  ODSClearAll;
  impODS.Free;
  {$ENDIF}
  {$ENDIF}

  {$IFDEF ODT}
  {$IFDEF VCL6}
  ODTClearAll;
  impODT.Free;
  {$ENDIF}
  {$ENDIF}

  {$IFDEF XMLDOC}
  {$IFDEF VCL6}
  tvXMLDoc.Items.Clear;
  XMLDocClearList;
  FXMLDocFile.Free;
  impXMLDoc.Free;
  {$ENDIF}
  {$ENDIF}

  XLSClearFieldList;
  TXTClearCombo;
  CSVClearCombo;
  DBFClearList;
  DBFClearTableList;
  XMLClearList;
  XMLClearTableList;
  {$IFDEF ADO}
  AccessClearList;
  {$ENDIF}
  FormatsClearList;

  for i := 0 to pgImport.PageCount - 1 do
  begin
    pgImport.Pages[i].Name := pgImport.Pages[i].Name;
    pgImport.Pages[i].Parent := pgImport;
  end;
end;

procedure TQImport3WizardF.FormShow(Sender: TObject);
begin
  SetCaptions;
  SetTitle;
end;

procedure TQImport3WizardF.rbtClick(Sender: TObject);
var
  Rbt: TRadioButton;
begin
  if not (Sender is TRadioButton) then Exit;
  Rbt := (Sender as TRadioButton);
  if Rbt.Name = 'rbtXLS' then ImportType := aiXLS
  else if Rbt.Name = 'rbtXlsx' then ImportType := aiXlsx
  else if Rbt.Name = 'rbtDocx' then ImportType := aiDocx
  else if Rbt.Name = 'rbtODS' then ImportType := aiODS
  else if Rbt.Name = 'rbtODT' then ImportType := aiODT
  else if Rbt.Name = 'rbtDBF' then ImportType := aiDBF
  else if Rbt.Name = 'rbtHTML' then ImportType := aiHTML
  else if Rbt.Name = 'rbtXML' then ImportType := aiXML
  else if Rbt.Name = 'rbtXMLDoc' then ImportType := aiXMLDoc
  else if Rbt.Name = 'rbtTXT' then ImportType := aiTXT
  else if Rbt.Name = 'rbtCSV' then ImportType := aiCSV
  else if Rbt.Name = 'rbtAccess' then ImportType := aiAccess
end;

procedure TQImport3WizardF.cbCommaExit(Sender: TObject);
begin
  if AnsiCompareText(cbComma.Text, 'TAB') = 0 then
    Comma := Chr(VK_TAB)
  else if AnsiCompareText(cbComma.Text, 'SPACE') = 0 then
    Comma := Chr(VK_SPACE)
  else Comma := Str2Char(cbComma.Text,Comma);
end;

procedure TQImport3WizardF.cbQuoteExit(Sender: TObject);
begin
  Quote := Str2Char(cbQuote.Text,Quote);
end;

procedure TQImport3WizardF.cbXMLDocDataLocationChange(Sender: TObject);
begin
{$IFDEF XMLDOC}
{$IFDEF VCL6}
  XMLDocDataLocation := TXMLDataLocation(cbXMLDocDataLocation.ItemIndex);
{$ENDIF}
{$ENDIF}
end;

procedure TQImport3WizardF.edtFileNameChange(Sender: TObject);
begin
  FileName := edtFileName.Text
end;

procedure TQImport3WizardF.chGoToLastPageClick(Sender: TObject);
begin
  GoToLastPage := chGoToLastPage.Checked;
  Wizard.GoToLastPage := GoToLastPage;
end;

procedure TQImport3WizardF.chAutoSaveTemplateClick(Sender: TObject);
begin
  AutoSaveTemplate := chAutoSaveTemplate.Checked;
  Wizard.AutoSaveTemplate := AutoSaveTemplate;
end;

procedure TQImport3WizardF.chAutoTrimValueClick(Sender: TObject);
begin
  AutoTrimValue := chAutoTrimValue.Checked;
end;

procedure TQImport3WizardF.spbBrowseClick(Sender: TObject);
begin
  if opnDialog.Execute then edtFileName.Text := opnDialog.FileName;
end;

procedure TQImport3WizardF.SetImportType(const Value: TAllowedImport);
begin
  if FImportType <> Value then
  begin
    if not (Value in Wizard.AllowedImports) then
      raise EQImportError.Create(QImportLoadStr(QIE_UnknownImportSource));

    FImportType := Value;
    //----
    rbtXLS.Checked := FImportType = aiXLS;
    rbtXlsx.Checked := FImportType = aiXlsx;
    rbtDocx.Checked := FImportType = aiDocx;
    rbtODS.Checked := FImportType = aiODS;
    rbtODT.Checked := FImportType = aiODT;
    rbtDBF.Checked := FImportType = aiDBF;
    rbtHTML.Checked := FImportType = aiHTML;
    rbtXML.Checked := FImportType = aiXML;
    cbXMLDocumentType.Enabled  := FImportType = aiXML;
    laXMLDocumentType.Enabled  := FImportType = aiXML;
    rbtXMLDoc.Checked := FImportType = aiXMLDoc;
    rbtTXT.Checked := FImportType = aiTXT;
    rbtCSV.Checked := FImportType = aiCSV;
    rbtAccess.Checked := FImportType = aiAccess;
    //----
    TuneOpenDialog;
    ChangeExtension;
    {$IFDEF ADO}
    TuneAccessControls;
    {$ENDIF}
    TuneCSVControls;
    TuneXMLControls;
    if FImportType = aiCSV then
    begin
      if AnsiCompareText(cbComma.Text, 'TAB') = 0 then
        Comma := Chr(VK_TAB)
      else if AnsiCompareText(cbComma.Text, 'SPACE') = 0 then
        Comma := Chr(VK_SPACE)
      else
        Comma := Str2Char(cbComma.Text, Comma);
      Quote := Str2Char(cbQuote.Text, Quote);
    end
    else
    begin
      FComma := #0;
      FQuote := #0;
    end;
    TuneButtons;
  end;
end;

procedure TQImport3WizardF.SetFileName(const Value: qiString);
var
  i: Integer;
begin
  if AnsiCompareText(FFileName, Trim(Value)) <> 0 then
  begin
    FFileName := Trim(Value);
    edtFileName.Text := FFileName;
    //----
    case ImportType of
      aiXLS: begin
        lvXLSRanges.Items.Clear;
        for i := 0 to lvXLSFields.Items.Count - 1 do
          TMapRow(lvXLSFields.Items[i].Data).AsString := EmptyStr;
        XLSClearDataSheets;
        XLSSkipCols := 0;
        XLSSkipRows := 0;
      end;
    end;
    //----
    TuneButtons;
    SetTitle;
    FNeedLoadFile := True;
  end;
end;

procedure TQImport3WizardF.SetGoToLastPage(Value: boolean);
begin
  if FGoToLastPage <> Value then begin
    FGoToLastPage := Value;
    chGoToLastPage.Checked := FGoToLastPage;
  end;
end;

procedure TQImport3WizardF.SetAutoSaveTemplate(Value: boolean);
begin
  if FAutoSaveTemplate <> Value then begin
    FAutoSaveTemplate := Value;
    chAutoSaveTemplate.Checked := FAutoSaveTemplate;
  end;
end;

procedure TQImport3WizardF.SetAutoTrimValue(const Value: Boolean);
begin
  if FAutoTrimValue <> Value then
  begin
    FAutoTrimValue := Value;
    chAutoTrimValue.Checked := FAutoTrimValue;
  end;
end;

procedure TQImport3WizardF.SetComma(const Value: AnsiChar);
var
  i: Integer;
begin
  if FComma <> Value then begin
    FComma := Value;
    if FComma = Chr(VK_TAB) then
      cbComma.Text := 'TAB'
    else if FComma = Chr(VK_SPACE) then
      cbComma.Text := 'SPACE'
    else cbComma.Text := Char2Str(FComma);

    for i := 0 to lvCSVFields.Items.Count - 1 do
      lvCSVFields.Items[i].SubItems[0] := EmptyStr;
    if lvCSVFields.Items.Count > 0 then begin
      lvCSVFields.Items[0].Focused := True;
      lvCSVFields.Items[0].Selected := True;
    end;

    FNeedLoadFile := True;
    TuneButtons;
  end;
end;

procedure TQImport3WizardF.SetQuote(const Value: AnsiChar);
var
  i: Integer;
begin
  if FQuote <> Value then begin
    FQuote := Value;
    cbQuote.Text := Char2Str(FQuote);

    for i := 0 to lvCSVFields.Items.Count - 1 do
      lvCSVFields.Items[i].SubItems[0] := EmptyStr;
    if lvCSVFields.Items.Count > 0 then begin
      lvCSVFields.Items[0].Focused := True;
      lvCSVFields.Items[0].Selected := True;
    end;

    FNeedLoadFile := True;
  end;
end;

procedure TQImport3WizardF.FillCombosAndLists;
begin
  XLSFillFieldList;
  TXTFillCombo;
  CSVFillCombo;
  DBFFillList;
  {$IFDEF XLSX}
  {$IFDEF VCL6}
  XlsxFillList;{$ENDIF}{$ENDIF}
  {$IFDEF DOCX}
  {$IFDEF VCL6}
  DocxFillList;{$ENDIF}{$ENDIF}
  {$IFDEF ODS}
  {$IFDEF VCL6}
  ODSFillList;{$ENDIF}{$ENDIF}
  {$IFDEF ODT}
  {$IFDEF VCL6}
  ODTFillList;{$ENDIF}{$ENDIF}
  {$IFDEF HTML}
  {$IFDEF VCL6}
  HTMLFillList;{$ENDIF}{$ENDIF}
  XMLFillList;
  {$IFDEF XMLDOC}
  {$IFDEF VCL6}
  XMLDocFillList;{$ENDIF}{$ENDIF}
  {$IFDEF ADO}
  AccessFillList;{$ENDIF}
  FormatsFillList;
  FillKeyColumns(KeyColumns);
end;

procedure TQImport3WizardF.FillKeyColumns(Strings: TStrings);
var
  i, j: Integer;
  str: string;
begin
  lvAvailableColumns.Items.BeginUpdate;
  try
    lvSelectedColumns.Items.BeginUpdate;
    try
      lvAvailableColumns.Items.Clear;
      lvSelectedColumns.Items.Clear;

      for i := 0 to Strings.Count - 1 do begin
        j := QImportDestinationFindColumn(False, ImportDestination, DataSet,
             DBGrid, ListView, StringGrid, GridCaptionRow, Strings[i]);
        if j > -1 then
          lvSelectedColumns.Items.Add.Caption :=
            QImportDestinationColName(False, ImportDestination, DataSet,
              DBGrid, ListView, StringGrid, GridCaptionRow, j);
      end;

      for i := 0 to QImportDestinationColCount(False, ImportDestination,
                      DataSet, DBGrid, ListView, StringGrid) - 1 do begin
         str := QImportDestinationColName(False, ImportDestination,
          DataSet, DBGrid, ListView, StringGrid, GridCaptionRow, i);
         if Strings.IndexOf(str) = -1 then
           lvAvailableColumns.Items.Add.Caption := str;
      end;
      if lvAvailableColumns.Items.Count > 0 then begin
        lvAvailableColumns.Items[0].Focused := True;
        lvAvailableColumns.Items[0].Selected := True;
      end;
      if lvSelectedColumns.Items.Count > 0 then begin
        lvSelectedColumns.Items[0].Focused := True;
        lvSelectedColumns.Items[0].Selected := True;
      end;
    finally
      lvSelectedColumns.Items.EndUpdate;
    end;
  finally
    lvAvailableColumns.Items.EndUpdate;
  end;
end;

procedure TQImport3WizardF.MoveToSelected(Source, Destination: TListView;
  All: boolean; Index: Integer);
var
  List: TStringList;
  i: Integer;
  ListItem: TListItem;
begin
  Source.Items.BeginUpdate;
  try
    Destination.Items.BeginUpdate;
    try
      List := TStringList.Create;
      try
        for i := Source.Items.Count - 1 downto 0 do
          if Source.Items[i].Selected or All then begin
            List.Add(Source.Items[i].Caption);
            Source.Items.Delete(i);
          end;
        ListItem := nil;
        if (List.Count = 1) and (Index > -1) then begin
          ListItem := Destination.Items.Insert(Index);
          ListItem.Caption := List[0];
          List.Delete(0);
        end
        else
          for i := List.Count - 1 downto 0 do begin
            ListItem := Destination.Items.Add;
            ListItem.Caption := List[i];
            List.Delete(i);
          end;
      finally
        List.Free;
      end;
      if Assigned(Source.ItemFocused) then
        Source.ItemFocused.Selected := True;
      if Assigned(ListItem) then
        for i := 0 to Destination.Items.Count - 1 do
          Destination.Items[i].Selected := Destination.Items[i] = ListItem;
    finally
      Destination.Items.EndUpdate;
    end;
  finally
    Source.Items.EndUpdate;
  end;
end;

procedure TQImport3WizardF.SetCaptions;
var
  N: Integer;
  I: TQICharsetType;
begin
  grpImportTypes.Caption := QImportLoadStr(QIW_ImportFrom);
  rbtXLS.Caption := QImportLoadStr(QIW_XLS);
  rbtXlsx.Caption := QImportLoadStr(QIW_XLSX);
  rbtDocx.Caption := QImportLoadStr(QIW_DOCX);
  rbtODS.Caption := QImportLoadStr(QIW_ODS);
  rbtODT.Caption := QImportLoadStr(QIW_ODT);  
  rbtDBF.Caption := QImportLoadStr(QIW_DBF);
  rbtHTML.Caption := QImportLoadStr(QIW_HTML);
  rbtAccess.Caption := QImportLoadStr(QIW_Access);
  laAccessPassword.Caption := QImportLoadStr(QIW_Access_Password);
  rbtXML.Caption := QImportLoadStr(QIW_XML);
  rbtXMLDoc.Caption := QImportLoadStr(QIW_XMLDoc);
  rbtTXT.Caption := QImportLoadStr(QIW_TXT);
  rbtCSV.Caption := QImportLoadStr(QIW_CSV);
  laComma.Caption := QImportLoadStr(QIW_Comma);
  laQuote.Caption := QImportLoadStr(QIW_Quote);
  laSourceFileName.Caption := QImportLoadStr(QIW_FileName);
  laTemplateFileNameTag.Caption := QImportLoadStr(QIW_TemplateFileName);
  odTemplate.Filter := QImportLoadStr(QIF_TEMPLATE);
  sdTemplate.Filter := QImportLoadStr(QIF_TEMPLATE);
  bHelp.Caption := QImportLoadStr(QIW_Help);
  bBack.Caption := QImportLoadStr(QIW_Back);
  bNext.Caption := QImportLoadStr(QIW_Next);
  bCancel.Caption := QImportLoadStr(QIW_Cancel);
  bOk.Caption := QImportLoadStr(QIW_Execute);
  gbTemplateOptions.Caption := QImportLoadStr(QIW_TemplateOptions);
  chGoToLastPage.Caption := QImportLoadStr(QIW_GoToLastPage);
  chAutoSaveTemplate.Caption := QImportLoadStr(QIW_AutoSaveTemplate);
  btnLoadTemplate.Caption := QImportLoadStr(QIW_LoadTemplate);
  //-----
  laTXTEncoding.Caption := QImportLoadStr(QIW_TXT_ENCODING);
  lvTXTFields.Columns[0].Caption := QImportLoadStr(QIW_TXT_Fields);
  lvTXTFields.Columns[1].Caption := QImportLoadStr(QIW_TXT_Fields_Pos);
  lvTXTFields.Columns[2].Caption := QImportLoadStr(QIW_TXT_Fields_Size);
  tbtTXTClear.Hint := QImportLoadStr(QIW_TXT_Clear);
  laTXTSkipLines.Caption := QImportLoadStr(QIW_TXT_SkipLines);
  for I := Low(TQICharsetType) to High(TQICharsetType) do
    cmbTXTEncoding.Items.Add(QICharsetTypeNames[I]);
  //-----
  lvCSVFields.Columns[0].Caption := QImportLoadStr(QIW_CSV_Fields);
  laCSVColNumber.Caption := QImportLoadStr(QIW_CSV_ColNumber);
  laCSVSkipLines.Caption := QImportLoadStr(QIW_CSV_SkipLines);
  tbtCSVAutoFill.Hint := QImportLoadStr(QIW_CSV_AutoFill);
  tbtCSVClear.Hint := QImportLoadStr(QIW_CSV_Clear);
  laCSVEncoding.Caption := QImportLoadStr(QIW_TXT_ENCODING);
  for I := Low(TQICharsetType) to High(TQICharsetType) do
    cmbCSVEncoding.Items.Add(QICharsetTypeNames[I]);
  //-----
  bDBFAdd.Caption := QImportLoadStr(QIW_DBF_Add);
  bDBFAutoFill.Caption := QImportLoadStr(QIW_DBF_AutoFill);
  bDBFRemove.Caption := QImportLoadStr(QIW_DBF_Remove);
  bDBFClear.Caption := QImportLoadStr(QIW_DBF_Clear);
  chDBFSkipDeleted.Caption := QImportLoadStr(QIW_DBF_SkipDeleted);
  //-----
  rbtAccessTable.Caption := QImportLoadStr(QIW_Access_Table);
  rbtAccessSQL.Caption := QImportLoadStr(QIW_Access_SQL);
  tbtAccessSQLSave.Hint := QImportLoadStr(QIW_Access_SQL_Save);
  tbtAccessSQLLoad.Hint := QImportLoadStr(QIW_Access_SQL_Load);
  bAccessAdd.Caption := QImportLoadStr(QIW_Access_Add);
  bAccessAutoFill.Caption := QImportLoadStr(QIW_Access_AutoFill);
  bAccessRemove.Caption := QImportLoadStr(QIW_Access_Remove);
  bAccessClear.Caption := QImportLoadStr(QIW_Access_Clear);
  //-----
  lvXlsxFields.Columns[0].Caption := QImportLoadStr(QIW_Xlsx_Fields);
  //dee
  lvXlsxFields.Columns[1].Caption := '';//QImportLoadStr(QIW_Xlsx_Col);
  //\
  laXlsxSkipRows.Caption := QImportLoadStr(QIW_Xlsx_SkipRows);
  tlbXlsxAutoFill.Hint := QImportLoadStr(QIW_Xlsx_AutoFill);
  tlbXlsxClear.Hint := QImportLoadStr(QIW_Xlsx_Clear);
  //-----
  lvDocxFields.Columns[0].Caption := QImportLoadStr(QIW_Docx_Fields);
  lvDocxFields.Columns[1].Caption := ''; //dee
  laDocxSkipRows.Caption := QImportLoadStr(QIW_Docx_SkipRows);
  tlbDocxAutoFill.Hint := QImportLoadStr(QIW_Docx_AutoFill);
  tlbDocxClear.Hint := QImportLoadStr(QIW_Docx_Clear);
  //-----
  lvODSFields.Columns[0].Caption := QImportLoadStr(QIW_ODS_Fields);
  lvODSFields.Columns[1].Caption := ''; //dee
  laODSSkipRows.Caption := QImportLoadStr(QIW_ODS_SkipRows);
  tlbODSAutoFill.Hint := QImportLoadStr(QIW_ODS_AutoFill);
  tlbODSClear.Hint := QImportLoadStr(QIW_ODS_Clear);
  //-----
  lvODTFields.Columns[0].Caption := QImportLoadStr(QIW_ODT_Fields);
  lvODTFields.Columns[1].Caption := ''; //dee
  laODTSkipRows.Caption := QImportLoadStr(QIW_ODT_SkipRows);
  tlbODTAutoFill.Hint := QImportLoadStr(QIW_ODT_AutoFill);
  tlbODTClear.Hint := QImportLoadStr(QIW_ODT_Clear);
  cbODTUseHeader.Caption := QImportLoadStr(QIW_ODT_UseFirstRowAsHeader);
  //-----
  lvHTMLFields.Columns[0].Caption := QImportLoadStr(QIW_HTML_Fields);
  lvHTMLFields.Columns[1].Caption := ''; //dee
  laHTMLTableNumber.Caption := QImportLoadStr(QIW_HTML_TableNumber);
  laHTMLColNumber.Caption := QImportLoadStr(QIW_HTML_ColNumber);
  laHTMLSkipLines.Caption := QImportLoadStr(QIW_HTML_SkipLines);
  tbtHTMLAutoFill.Hint := QImportLoadStr(QIW_HTML_AutoFill);
  tbtHTMLClear.Hint := QImportLoadStr(QIW_HTML_Clear);
  //-----
  bXMLAdd.Caption := QImportLoadStr(QIW_XML_Add);
  bXMLAutoFill.Caption := QImportLoadStr(QIW_XML_AutoFill);
  bXMLRemove.Caption := QImportLoadStr(QIW_XML_Remove);
  bXMLClear.Caption := QImportLoadStr(QIW_XML_Clear);
  chXMLWriteOnFly.Caption := QImportLoadStr(QIW_XML_WriteOnFly);
  //-----
  lvXMLDocFields.Columns[0].Caption := QImportLoadStr(QIW_XMLDoc_Fields);
  lvXMLDocFields.Columns[1].Caption := ''; //dee
  laXMLDocColNumber.Caption := QImportLoadStr(QIW_XMLDoc_ColNumber);
  laXMLDocSkipLines.Caption := QImportLoadStr(QIW_XMLDoc_SkipLines);
  tbtXMLDocAutoFill.Hint := QImportLoadStr(QIW_XMLDoc_AutoFill);
  tbtXMLDocClear.Hint := QImportLoadStr(QIW_XMLDoc_Clear);
  laXMLDocXPath.Caption := QImportLoadStr(QIW_XMLDoc_XPath);
  laXMLDocDataLocation.Caption := QImportLoadStr(QIW_XMLDoc_DataLocation);
  bXMLDocBuildTree.Caption := QImportLoadStr(QIW_XMLDoc_BuildTreeView);
  bXMLDocGetXPath.Caption := QImportLoadStr(QIW_XMLDoc_GetXPath);

  //-----
  lvXLSFields.Columns[0].Caption := QImportLoadStr(QIW_XLS_Fields);
  lvXLSRanges.Columns[0].Caption := QImportLoadStr(QIW_XLS_Ranges);
  //dee
  lvXLSSelection.Columns[0].Caption := QImportLoadStr(QIW_XLS_Ranges);
  //\
  tbtXLSAddRange.Hint := QImportLoadStr(QIW_XLS_AddRange);
  tbtXLSEditRange.Hint := QImportLoadStr(QIW_XLS_EditRange);
  tbtXLSDelRange.Hint := QImportLoadStr(QIW_XLS_DelRange);
  tbtXLSMoveRangeUp.Hint := QImportLoadStr(QIW_XLS_MoveRangeUp);
  tbtXLSMoveRangeDown.Hint := QImportLoadStr(QIW_XLS_MoveRangeDown);
  laXLSSkipCols.Caption := QImportLoadStr(QIW_XLS_SkipCols);
  laXLSSkipRows.Caption := QImportLoadStr(QIW_XLS_SkipRows);
  tbtXLSAutoFillCols.Hint := QImportLoadStr(QIW_XLS_AutoFillCols);
  tbtXLSAutoFillRows.Hint := QImportLoadStr(QIW_XLS_AutoFillRows);
  tbtXLSClearFieldRanges.Hint := QImportLoadStr(QIW_XLS_ClearFieldRanges);
  tbtXLSClearAllRanges.Hint := QImportLoadStr(QIW_XLS_ClearAllRanges);
  //-----
  tshBaseFormats.Caption := QImportLoadStr(QIW_BaseFormats);
  grpDateTimeFormats.Caption := QImportLoadStr(QIW_DateTimeFormats);
  grpSeparators.Caption := QImportLoadStr(QIW_Separators);
  laDecimalSeparator.Caption := QImportLoadStr(QIW_DecimalSeparator);
  laThousandSeparator.Caption := QImportLoadStr(QIW_ThousandSeparator);
  laShortDateFormat.Caption := QImportLoadStr(QIW_ShortDateFormat);
  laLongDateFormat.Caption := QImportLoadStr(QIW_LongDateFormat);
  laDateSeparator.Caption := QImportLoadStr(QIW_DateSeparator);
  laShortTimeFormat.Caption := QImportLoadStr(QIW_ShortTimeFormat);
  laLongTimeFormat.Caption := QImportLoadStr(QIW_LongTimeFormat);
  laTimeSeparator.Caption := QImportLoadStr(QIW_TimeSeparator);

  laBooleanTrue.Caption := QImportLoadStr(QIW_BooleanTrue);
  mmBooleanTrue.Lines.Assign(Wizard.Formats.BooleanTrue);
  laBooleanFalse.Caption := QImportLoadStr(QIW_BooleanFalse);
  mmBooleanFalse.Lines.Assign(Wizard.Formats.BooleanFalse);
  laNullValues.Caption := QImportLoadStr(QIW_NullValue);
  mmNullValues.Lines.Assign(Wizard.Formats.NullValues);
  //-----------------------
  tshDataFormats.Caption := QImportLoadStr(QIWDF_Caption);
  lstFormatFields.Columns[0].Caption := QImportLoadStr(QIWDF_Fields);
  tsFieldTuning.Caption := QImportLoadStr(QIWDF_Tuning);
  laGeneratorValue.Caption := QImportLoadStr(QIWDF_GeneratorValue);
  laGeneratorStep.Caption := QImportLoadStr(QIWDF_GeneratorStep);
  laConstantValue.Caption := QImportLoadStr(QIWDF_ConstantValue);
  laNullValue.Caption := QImportLoadStr(QIWDF_NullValue);
  laDefaultValue.Caption := QImportLoadStr(QIWDF_DefaultValue);
  laLeftQuote.Caption := QImportLoadStr(QIWDF_LeftQuote);
  laRightQuote.Caption := QImportLoadStr(QIWDF_RightQuote);
  laQuoteAction.Caption := QImportLoadStr(QIWDF_QuoteAction);
  N := cmbQuoteAction.ItemIndex;
  cmbQuoteAction.Items[0] := QImportLoadStr(QIWDF_QuoteNone);
  cmbQuoteAction.Items[1] := QImportLoadStr(QIWDF_QuoteAdd);
  cmbQuoteAction.Items[2] := QImportLoadStr(QIWDF_QuoteRemove);
  cmbQuoteAction.ItemIndex := N;
  laCharCase.Caption := QImportLoadStr(QIWDF_CharCase);
  laCharSet.Caption := QImportLoadStr(QIWDF_CharSet);
  N := cmbCharCase.ItemIndex;
  cmbCharCase.Items[0] := QImportLoadStr(QIWDF_CharCaseNone);
  cmbCharCase.Items[1] := QImportLoadStr(QIWDF_CharCaseUpper);
  cmbCharCase.Items[2] := QImportLoadStr(QIWDF_CharCaseLower);
  cmbCharCase.Items[3] := QImportLoadStr(QIWDF_CharCaseUpperFirst);
  cmbCharCase.Items[4] := QImportLoadStr(QIWDF_CharCaseUpperWord);
  cmbCharCase.ItemIndex := N;
  N := cmbCharSet.ItemIndex;
  cmbCharSet.Items[0] := QImportLoadStr(QIWDF_CharSetNone);
  cmbCharSet.Items[1] := QImportLoadStr(QIWDF_CharSetAnsi);
  cmbCharSet.Items[2] := QImportLoadStr(QIWDF_CharSetOem);
  cmbCharSet.ItemIndex := N;
  laReplacements.Caption := QImportLoadStr(QIWDF_Replacements);
  lScript.Caption := QImportLoadStr(QIWDF_JScript);
  lvReplacements.Columns[0].Caption := QImportLoadStr(QIWDF_TextToFind);
  lvReplacements.Columns[1].Caption := QImportLoadStr(QIWDF_ReplaceWith);
  lvReplacements.Columns[2].Caption := QImportLoadStr(QIWDF_IgnoreCase);
  tbtAddReplacement.Hint := QImportLoadStr(QIWDF_AddReplacement);
  tbtEditReplacement.Hint := QImportLoadStr(QIWDF_EditReplacement);
  tbtDelReplacement.Hint := QImportLoadStr(QIWDF_DelReplacement);
  //-----------------------
  tshCommit.Caption := QImportLoadStr(QIW_CommitOptions);
  grpCommit.Caption := QImportLoadStr(QIW_Commit);
  chCommitAfterDone.Caption := QImportLoadStr(QIW_CommitAfterDone);
  laCommitRecCount_01.Caption := QImportLoadStr(QIW_CommitRecCount);
  laCommitRecCount_02.Caption := QImportLoadStr(QIW_Records);

  grpImportCount.Caption := QImportLoadStr(QIW_RecordCount);
  chImportAllRecords.Caption := QImportLoadStr(QIW_ImportAllRecords);
  laImportRecCount_01.Caption := QImportLoadStr(QIW_ImportRecCount);
  laImportRecCount_02.Caption := QImportLoadStr(QIW_Records);
  chCloseAfterImport.Caption := QImportLoadStr(QIW_CloseAfterImport);
  chAutoTrimValue.Caption := QImportLoadStr(QIW_AutoTrimValue);
  chImportEmptyRows.Caption := QImportLoadStr(QIW_ImportEmptyRows);

  grpErrorLog.Caption := QImportLoadStr(QIW_ErrorLog);
  chEnableErrorLog.Caption := QImportLoadStr(QIW_EnableErrorLog);
  laErrorLogFileName.Caption := QImportLoadStr(QIW_ErrorLogFileName);
  chRewriteErrorLogFile.Caption := QImportLoadStr(QIW_RewriteErrorLogFile);
  chShowErrorLog.Caption := QImportLoadStr(QIW_ShowErrorLog);
  btnSaveTemplate.Caption := QImportLoadStr(QIW_SaveTemplate);

  tshAdvanced.Caption := QImportLoadStr(QIW_ImportAdvanced);
  rgAddType.Caption := QImportLoadStr(QIW_AddType);
  N := rgAddType.ItemIndex;
  rgAddType.Items[0] := QImportLoadStr(QIW_AddType_Append);
  rgAddType.Items[1] := QImportLoadStr(QIW_AddType_Insert);
  rgAddType.ItemIndex := N;

  rgImportMode.Caption := QImportLoadStr(QIWIM_Caption);
  N := rgImportMode.ItemIndex;
  rgImportMode.Items[0] := QImportLoadStr(QIWIM_Insert_All);
  rgImportMode.Items[1] := QImportLoadStr(QIWIM_Insert_New);
  rgImportMode.Items[2] := QImportLoadStr(QIWIM_Update);
  rgImportMode.Items[3] := QImportLoadStr(QIWIM_Update_or_Insert);
  rgImportMode.Items[4] := QImportLoadStr(QIWIM_Delete);
  rgImportMode.Items[5] := QImportLoadStr(QIWIM_Delete_or_Insert);
  rgImportMode.ItemIndex := N;

  laAvailableColumns.Caption := QImportLoadStr(QIW_AvailableColumns);
  laSelectedColumns.Caption := QImportLoadStr(QIW_SelectedColumns);
end;

procedure TQImport3WizardF.TuneVisible;
begin
  {!TODO}
{  if not (aiHTML in Wizard.AllowedImports) then
  begin
    rbtHTML.Visible := False;
    Self.ClientHeight := Self.ClientHeight - 30;
    rbtXMLDoc.Top := rbtHTML.Top;
    rbtAccess.Top := rbtXMLDoc.Top + 27;
    laAccessPassword.Top := rbtAccess.Top + 20;
    edAccessPassword.Top := rbtAccess.Top + 20;
  end;
  if not (aiXMLDoc in Wizard.AllowedImports) then
  begin
    rbtXMLDoc.Visible := False;
    Self.ClientHeight := Self.ClientHeight - 30;
    rbtAccess.Top := rbtXMLDoc.Top;
    laAccessPassword.Top := rbtAccess.Top + 20;
    edAccessPassword.Top := rbtAccess.Top + 20;
  end;
  if not (aiAccess in Wizard.AllowedImports) then
  begin
    rbtAccess.Visible := False;
    laAccessPassword.Visible := False;
    edAccessPassword.Visible := False;
    Self.ClientHeight := Self.ClientHeight - 50;
  end;}
end;

procedure TQImport3WizardF.TuneXMLControls;
begin
  laXMLDocumentType.Enabled := FImportType = aiXML;
  cbXMLDocumentType.Enabled := FImportType = aiXML;
  if FImportType = aiXML then
    cbXMLDocumentType.ItemIndex := 0;
end;

procedure TQImport3WizardF.TuneOpenDialog;
begin
  case FImportType of
    aiXLS: opnDialog.Filter := QImportLoadStr(QIF_XLS);
    aiXlsx: opnDialog.Filter := QImportLoadStr(QIF_XLSX);
    aiDocx: opnDialog.Filter := QImportLoadStr(QIF_DOCX);
    aiODS: opnDialog.Filter := QImportLoadStr(QIF_ODS);
    aiODT: opnDialog.Filter := QImportLoadStr(QIF_ODT);
    aiDBF: opnDialog.Filter := QImportLoadStr(QIF_DBF);
    aiHTML: opnDialog.Filter := QImportLoadStr(QIF_HTML);
    aiXML: opnDialog.Filter := QImportLoadStr(QIF_XML);
    aiTXT: opnDialog.Filter := QImportLoadStr(QIF_TXT);
    aiCSV: opnDialog.Filter := QImportLoadStr(QIF_CSV);
    aiXMLDoc: opnDialog.Filter := QImportLoadStr(QIF_XML);
    aiAccess: opnDialog.Filter := QImportLoadStr(QIF_Access);
  end;
end;

procedure TQImport3WizardF.ChangeExtension;
begin
  if not Wizard.AutoChangeExtension then Exit;
  if Trim(FileName) = EmptyStr then Exit;
  FileName := ChangeFileExt(FileName, FileExts[Ord(FImportType)]);
end;

procedure TQImport3WizardF.TuneCSVControls;
begin
  laComma.Enabled := FImportType = aiCSV;
  cbComma.Enabled := FImportType = aiCSV;
  if cbComma.Enabled  and (cbComma.Text = EmptyStr) then
    Comma := AnsiChar(GetListSeparator);
  laQuote.Enabled := FImportType = aiCSV;
  cbQuote.Enabled := FImportType = aiCSV;
  if cbQuote.Enabled  and (cbQuote.Text = EmptyStr) then
    Quote := '"';
end;

function TQImport3WizardF.GetWizard: TQImport3Wizard;
begin
  Result := Owner as TQImport3Wizard;
end;

function TQImport3WizardF.GetTemplateFileName: qiString;
begin
  Result := Wizard.TemplateFileName;
end;

function TQImport3WizardF.GetAutoLoadTemplate: boolean;
begin
  Result := Wizard.AutoLoadTemplate;
end;

function TQImport3WizardF.GetImportDestination: TQImportDestination;
begin
  Result := Wizard.ImportDestination;
end;

function TQImport3WizardF.GetGridCaptionRow: Integer;
begin
  Result := Wizard.GridCaptionRow;
end;

function TQImport3WizardF.GetGridStartRow: Integer;
begin
  Result := Wizard.GridStartRow;
end;

function TQImport3WizardF.GetKeyColumns: TStrings;
begin
  Result := Wizard.KeyColumns;
end;

procedure TQImport3WizardF.SetStep(const Value: Integer);
begin
  if FStep <> Value then
  begin
    if Value = 1 then
    begin
      if not FileExists(FFileName) then
        raise EQImportError.CreateFmt(QImportLoadStr(QIE_FileNotExists), [FFileName]);

      if not AllowedImportFileType(FFileName) then
        raise EQImportError.CreateFmt(QImportLoadStr(QIE_WrongFileType), [ImportTypeStr(ImportType)]);

      if not ImportTypeEquivFileType(FFileName) then
        raise EQImportError.CreateFmt(QImportLoadStr(QIE_WrongFileType), [ImportTypeStr(ImportType)]);

      if not QImportIsDestinationActive(False, ImportDestination, DataSet,
               DBGrid, ListView, StringGrid) then
        raise EQImportError.Create(QImportLoadStr(QIE_DataSetClosed));
    end;
    FStep := Value;
    case Value of
      0: TuneStart;
      1: try
        case FImportType of
          aiXLS: XLSTune;
          aiDBF: DBFTune;
          {$IFDEF XLSX}
          {$IFDEF VCL6}
          aiXlsx: XlsxTune;
          {$ENDIF}{$ENDIF}

          {$IFDEF DOCX}
          {$IFDEF VCL6}
          aiDocx: DocxTune;
          {$ENDIF}{$ENDIF}

          {$IFDEF ODS}
          {$IFDEF VCL6}
          aiODS: ODSTune;
          {$ENDIF}{$ENDIF}

          {$IFDEF ODT}
          {$IFDEF VCL6}
          aiODT: ODTTune;
          {$ENDIF}{$ENDIF}

          {$IFDEF HTML}
          {$IFDEF VCL6}
          aiHTML: HTMLTune;
          {$ENDIF}{$ENDIF}

          aiXML: XMLTune;

          {$IFDEF XMLDOC}
          {$IFDEF VCL6}
          aiXMLDoc: XMLDocTune;
          {$ENDIF}{$ENDIF}
          
          aiTXT: TXTTune;
          aiCSV: CSVTune;
          {$IFDEF ADO}
          aiAccess: AccessTune_01;
          {$ENDIF}
        end;
      except
        Step := 0;
        raise;
      end;
      2: {$IFDEF ADO}if FImportType = aiAccess then
      begin
           try
             AccessTune_02;
           except
             Step := 1;
             raise;
           end;
         end
         else {$ENDIF}TuneFormats;
      3: if FImportType = aiAccess then
           TuneFormats
         else
           TuneFinish;
      4: if FImportType = aiAccess then
           TuneFinish;
    end;
  end;
end;

procedure TQImport3WizardF.ShowTip(Parent: TWinControl; Left, Top, Height,
  Width: Integer; const Tip: qiString);
begin
  paTip.Parent := Parent;
  paTip.Left := Left;
  paTip.Top := Top;
  paTip.Height := Height;
  paTip.Width := Width;
  paTip.Caption := Tip;
end;

procedure TQImport3WizardF.bNextClick(Sender: TObject);
begin
{$IFDEF ODT}
{$IFDEF VCL6}
  if (FImportType = aiODT) then
    if (Step = 1) then
      if (ODTComplex) then
        if MessageDlg('Selected table has a complex structure and could be improperly converted' + #10 +
          '(it contains vertically merged cells and/or subtables).'
            + #10 + 'Do you want to convert this table anyway?' + #10 +
              'You could examine internal structure of selected table by pressing No.',
                mtInformation, [mbYes, mbNo], 0) = mrNo then
    exit;
{$ENDIF}
{$ENDIF}
  Step := Step + 1;
end;

procedure TQImport3WizardF.bBackClick(Sender: TObject);
begin
  Step := Step - 1;
end;

procedure TQImport3WizardF.bDBFAddClick(Sender: TObject);
begin
  with lstDBFMap.Items.Add do begin
    Caption := lstDBFDataSet.Selected.Caption;
    SubItems.Add('=');
    SubItems.Add(lstDBF.Selected.Caption);
    Focused := True;
    Selected := True;
  end;
  lstDBFDataSet.Items.Delete(lstDBFDataSet.Selected.Index);
  if lstDBFDataSet.Items.Count > 0 then begin
    lstDBFDataSet.Items[0].Focused := True;
    lstDBFDataSet.Items[0].Selected := True;
  end;
  if lstDBFMap.Items.Count > 0 then begin
    lstDBFMap.Items[0].Focused := True;
    lstDBFMap.Items[0].Selected := True;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.bDBFAutoFillClick(Sender: TObject);
var
  i, N: Integer;
begin
  lstDBFDataSet.Items.BeginUpdate;
  try
    lstDBF.Items.BeginUpdate;
    try
      lstDBFMap.Items.BeginUpdate;
      try
        lstDBFMap.Items.Clear;
        DBFClearList;
        DBFFillList;
        N := lstDBF.Items.Count;
        if N > lstDBFDataSet.Items.Count
          then N := lstDBFDataSet.Items.Count;
        for i := N - 1 downto 0 do begin
          with lstDBFMap.Items.Insert(0) do begin
            Caption := lstDBFDataSet.Items[i].Caption;
            SubItems.Add('=');
            SubItems.Add(lstDBF.Items[i].Caption);
          end;
          lstDBFDataSet.Items[i].Delete;
        end;
        if lstDBFMap.Items.Count > 0 then begin
          lstDBFMap.Items[0].Focused := True;
          lstDBFMap.Items[0].Selected := True;
        end;
      finally
        lstDBFMap.Items.EndUpdate;
      end;
    finally
      lstDBF.Items.EndUpdate;
    end;
  finally
    lstDBFDataSet.Items.EndUpdate;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.bDBFRemoveClick(Sender: TObject);
begin
  lstDBFDataSet.Items.Add.Caption := lstDBFMap.Selected.Caption;
  lstDBFMap.Items.Delete(lstDBFMap.Selected.Index);
  if lstDBFMap.Items.Count > 0 then begin
    lstDBFMap.Items[0].Focused := True;
    lstDBFMap.Items[0].Selected := True;
  end;
  if lstDBFDataSet.Items.Count > 0 then begin
    lstDBFDataSet.Items[0].Focused := True;
    lstDBFDataSet.Items[0].Selected := True;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.bDBFClearClick(Sender: TObject);
begin
  lstDBFDataSet.Items.BeginUpdate;
  try
    lstDBFMap.Items.BeginUpdate;
    try
      lstDBFMap.Items.Clear;
      DBFClearList;
      DBFFillList;
    finally
      lstDBFMap.Items.EndUpdate;
    end;
  finally
    lstDBFDataSet.Items.EndUpdate;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (FImportType = aiXLS) and (Step = 1) then FShift := Shift;
end;

procedure TQImport3WizardF.FormKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (FImportType = aiXLS) and (Step = 1) then FShift := [];
end;

procedure TQImport3WizardF.TuneFinish;
begin
  pgImport.ActivePage := tsCommitOptions;
  if FImportType = aiAccess then
    laStep_02.Caption := Format(QImportLoadStr(QIW_Step), [4,4])
  else
    laStep_02.Caption := Format(QImportLoadStr(QIW_Step), [3,3]);
  bOk.Enabled := True;
  TuneButtons;
end;

procedure TQImport3WizardF.TuneButtons;
begin
  case Step of
    0: bNext.Enabled := StartReady;
    1: case FImportType of
         aiXLS: begin
           bNext.Enabled := XLSReady;
           tbtXLSAddRange.Enabled := Assigned(lvXLSFields.ItemFocused);
           tbtXLSEditRange.Enabled := Assigned(lvXLSFields.ItemFocused) and
                                      Assigned(lvXLSRanges.ItemFocused);
           tbtXLSDelRange.Enabled := Assigned(lvXLSFields.ItemFocused) and
                                     (lvXLSRanges.SelCount > 0);
           tbtXLSMoveRangeUp.Enabled := Assigned(lvXLSFields.ItemFocused) and
                                        Assigned(lvXLSRanges.ItemFocused) and
                                        (lvXLSRanges.ItemFocused.Index > 0);
           tbtXLSMoveRangeDown.Enabled := Assigned(lvXLSFields.ItemFocused) and
                                          Assigned(lvXLSRanges.ItemFocused) and
                                          (lvXLSRanges.ItemFocused.Index < lvXLSRanges.Items.Count - 1);
         end;
         aiDBF: begin
           bNext.Enabled := DBFReady;
           bDBFAdd.Enabled := Assigned(lstDBFDataSet.Selected) and
                              Assigned(lstDBF.Selected);
           bDBFRemove.Enabled := Assigned(lstDBFMap.Selected);
           bDBFClear.Enabled := lstDBFMap.Items.Count > 0;
         end;
         {$IFDEF XLSX}
         {$IFDEF VCL6}
         aiXlsx: bNext.Enabled := XlsxReady;
         {$ENDIF}{$ENDIF}

         {$IFDEF DOCX}
         {$IFDEF VCL6}
         aiDocx: bNext.Enabled := DocxReady;
         {$ENDIF}{$ENDIF}

         {$IFDEF ODS}
         {$IFDEF VCL6}
         aiODS: bNext.Enabled := ODSReady;
         {$ENDIF}{$ENDIF}

         {$IFDEF ODT}
         {$IFDEF VCL6}
         aiODT: bNext.Enabled := ODTReady;
         {$ENDIF}{$ENDIF}

         {$IFDEF HTML}
         {$IFDEF VCL6}
         aiHTML: bNext.Enabled := HTMLReady;
         {$ENDIF}{$ENDIF}

         aiXML: begin
           bNext.Enabled := XMLReady;
           bXMLAdd.Enabled := Assigned(lvXMLDataSet.Selected) and
                              Assigned(lvXML.Selected);
           bXMLRemove.Enabled := Assigned(lvXMLMap.Selected);
           bXMLClear.Enabled := lvXMLMap.Items.Count > 0;
         end;
         {$IFDEF XMLDOC}
         {$IFDEF VCL6}
         aiXMLDoc: bNext.Enabled := XMLDocReady;
         {$ENDIF}{$ENDIF}
         aiTXT: bNext.Enabled := TXTReady;
         aiCSV: bNext.Enabled := CSVReady;
         {$IFDEF ADO}
         aiAccess:
           bNext.Enabled := ((AccessSourceType = isTable) and
                             (lbAccessTables.ItemIndex > -1)) or
                            ((AccessSourceType = isSQL) and
                             (memAccessSQL.Text <> EmptyStr));
         {$ENDIF}
       end;
    2: {$IFDEF ADO}
      if FImportType = aiAccess then begin
        bNext.Enabled := AccessReady;
        bAccessAdd.Enabled := Assigned(lstAccessDataSet.Selected) and
                              Assigned(lstAccess.Selected);
        bAccessRemove.Enabled := Assigned(lstAccessMap.Selected);
        bAccessClear.Enabled := lstAccessMap.Items.Count > 0;
      end
      else{$ENDIF} begin
        bNext.Enabled := True;
        tbtAddReplacement.Enabled := Assigned(lstFormatFields.ItemFocused);
        tbtEditReplacement.Enabled := Assigned(lstFormatFields.ItemFocused) and
                                     Assigned(lvReplacements.ItemFocused) ;
        tbtDelReplacement.Enabled := Assigned(lstFormatFields.ItemFocused) and
                                     Assigned(lvReplacements.ItemFocused) ;
      end;
    3:
      begin
        bNext.Enabled := FImportType = aiAccess;
        if FImportType = aiAccess then begin
          tbtAddReplacement.Enabled := Assigned(lstFormatFields.ItemFocused);
          tbtEditReplacement.Enabled := Assigned(lstFormatFields.ItemFocused) and
                                        Assigned(lvReplacements.ItemFocused) ;
          tbtDelReplacement.Enabled := Assigned(lstFormatFields.ItemFocused) and
                                       Assigned(lvReplacements.ItemFocused) ;
      end;
    end;
    4: bNext.Enabled := false;
  end;
  bBack.Enabled := Step in [1,2,3,4];
  {$IFDEF ADO}
  if FImportType = aiAccess
    then bOk.Enabled := ((Step = 2) and AccessReady) or (Step > 2)
    else{$ENDIF} bOk.Enabled := ((Step = 1) and (((ImportType = aiXLS) and XLSReady) or
                                  ((ImportType = aiDBF) and DBFReady) or
                                  {$IFDEF XLSX}
                                  {$IFDEF VCL6}
                                  ((ImportType = aiXlsx) and XlsxReady) or
                                  {$ENDIF}{$ENDIF}

                                  {$IFDEF DOCX}
                                  {$IFDEF VCL6}
                                  ((ImportType = aiDocx) and DocxReady) or
                                  {$ENDIF}{$ENDIF}

                                  {$IFDEF ODS}
                                  {$IFDEF VCL6}
                                  ((ImportType = aiODS) and ODSReady) or
                                  {$ENDIF}{$ENDIF}

                                  {$IFDEF ODT}
                                  {$IFDEF VCL6}
                                  ((ImportType = aiODT) and ODTReady) or
                                  {$ENDIF}{$ENDIF}

                                  {$IFDEF HTML}
                                  {$IFDEF VCL6}
                                  ((ImportType = aiHTML) and HTMLReady) or
                                  {$ENDIF}{$ENDIF}

                                  ((ImportType = aiXML) and XMLReady) or
                                  {$IFDEF XMLDOC}
                                  {$IFDEF VCL6}
                                  ((ImportType = aiXMLDoc) and XMLDocReady) or
                                  {$ENDIF}{$ENDIF}

                                  ((ImportType = aiTXT) and TXTReady) or
                                  ((ImportType = aiCSV) and CSVReady))) or
                                  (Step > 1);
end;

function TQImport3WizardF.DBFReady: boolean;
begin
  Result := lstDBFMap.Items.Count > 0;
end;

function TQImport3WizardF.StartReady: boolean;
begin
  Result := ((FImportType = aiCSV) and (FComma <> #0) and
             (FFileName <> EmptyStr)) or
            ((FImportType <> aiCSV) and (FFileName <> EmptyStr));
end;

procedure TQImport3WizardF.TuneStart;
begin
  pgImport.ActivePage := tsImportType;
  TuneButtons;
end;

procedure TQImport3WizardF.bCancelClick(Sender: TObject);
begin
  Close;
end;

procedure TQImport3WizardF.bOkClick(Sender: TObject);
var
  i: Integer;
  fileDir: qiString;
begin
  case FImportType of
    aiXLS: Import := impXLS;
    aiDBF: Import := impDBF;
    {$IFDEF XLSX}
    {$IFDEF VCL6}
    aiXlsx: Import := impXlsx;
    {$ENDIF}
    {$ENDIF}

    {$IFDEF DOCX}
    {$IFDEF VCL6}
    aiDocx: Import := impDocx;
    {$ENDIF}
    {$ENDIF}

    {$IFDEF ODS}
    {$IFDEF VCL6}
    aiODS: Import := impODS;
    {$ENDIF}
    {$ENDIF}

    {$IFDEF ODT}
    {$IFDEF VCL6}
    aiODT: begin
      if (Step = 1) then
        if (ODTComplex) then
          if MessageDlg('Selected table has a complex structure and could be improperly converted' + #10 +
              '(it contains vertically merged cells and/or subtables).'
                + #10 + 'Do you want to convert this table anyway?' + #10 +
                'You could examine internal structure of selected table by pressing No.',
                  mtInformation, [mbYes, mbNo], 0) = mrNo then
            exit;
      Import := impODT;
    end;
    {$ENDIF}
    {$ENDIF}

    {$IFDEF HTML}
    {$IFDEF VCL6}
    aiHTML: Import := impHTML;
    {$ENDIF}
    {$ENDIF}
    aiXML: begin
            Import := impXML;
            TQImport3XML(Import).DocumentType := TQIXMLDocType(cbXMLDocumentType.ItemIndex);
           end;
    {$IFDEF XMLDOC}
    {$IFDEF VCL6}
    aiXMLDoc: Import := impXMLDoc;
    {$ENDIF}
    {$ENDIF}
    aiTXT,
    aiCSV: Import := impASCII;
    {$IFDEF ADO}
    aiAccess: Import := impAccess;
    {$ENDIF}
    else Import := nil;
  end;
  if Wizard.ShowProgress then
  begin
    FProgress := TfmQImport3ProgressDlg.CreateProgress(Self, Import);
    FProgress.Show;
    PostMessage(FProgress.Handle, WM_QIMPORT_PROGRESS, QIP_STATE, 0);
    Application.ProcessMessages;
  end;
  try
    Import.DataSet := DataSet;
    Import.DBGrid := DBGrid;
    Import.ListView := ListView;
    Import.StringGrid := StringGrid;
    Import.AutoTrimValue := AutoTrimValue;
    Import.ImportEmptyRows := ImportEmptyRows;
    {$IFDEF USESCRIPT}
    Import.ScriptEngine := FQImport3JScriptEngine;
    {$ENDIF}

    Import.ImportDestination := ImportDestination;
    Import.GridCaptionRow := GridCaptionRow;
    Import.GridStartRow := GridStartRow;

    Import.FileName := FileName;

    if Assigned(FProgress) then
      FProgress.Errors := Import.Errors;

    if Import is TQImport3DBF then
      with Import as TQImport3DBF do
      begin
        SkipDeleted := DBFSkipDeleted;
      end;

    {$IFDEF XLSX}
    {$IFDEF VCL6}
    if Import is TQImport3Xlsx then
      with Import as TQImport3Xlsx do
      begin
        SheetName := XlsxSheetName;
        SkipFirstRows := XlsxSkipLines;
      end;
    {$ENDIF}
    {$ENDIF}

    {$IFDEF DOCX}
    {$IFDEF VCL6}
    if Import is TQImport3Docx then
      with Import as TQImport3Docx do
      begin
        TableNumber := DocxTableNumber;
        SkipFirstRows := DocxSkipLines;
      end;
    {$ENDIF}
    {$ENDIF}

    {$IFDEF ODS}
    {$IFDEF VCL6}
    if Import is TQImport3ODS then
      with Import as TQImport3ODS do
      begin
        SheetName := ODSSheetName;
        SkipFirstRows := ODSSkipLines;
      end;
    {$ENDIF}
    {$ENDIF}

    {$IFDEF ODT}
    {$IFDEF VCL6}
    if Import is TQImport3ODT then
      with Import as TQImport3ODT do
      begin
        SheetName := ODTSheetName;
        SkipFirstRows := ODTSkipLines;
        UseHeader := ODTUseHeader;
      end;
    {$ENDIF}
    {$ENDIF}

    {$IFDEF HTML}
    {$IFDEF VCL6}
    if Import is TQImport3HTML then
      with Import as TQImport3HTML do
      begin
        TableNumber := Succ(cbHTMLTableNumber.ItemIndex);
        SkipFirstRows := HTMLSkipLines;
      end;
    {$ENDIF}
    {$ENDIF}

    if Import is TQImport3XML then
      with Import as TQImport3XML do
      begin
        WriteOnFly := XMLWriteOnFly;
      end;

    {$IFDEF XMLDOC}
    {$IFDEF VCL6}
    if Import is TQImport3XMLDoc then
      with Import as TQImport3XMLDoc do
      begin
        SkipFirstRows := XMLDocSkipLines;
        DataLocation := XMLDocDataLocation;
        XPath := XMLDocXPath;
      end;
    {$ENDIF}
    {$ENDIF}

    if Import is TQImport3ASCII then
      with Import as TQImport3ASCII do
      begin
        Quote := Self.Quote;
        Comma := Self.Comma;
        SkipFirstRows := TXTSkipLines * Integer(aiTXT = ImportType) +
                         CSVSkipLines * Integer(aiCSV = ImportType);
        if FImportType = aiTXT then
          Encoding := FTXTEncoding;

        if FImportType = aiCSV then
          Encoding := FCSVEncoding;
      end;

    if Import is TQImport3XLS then
      with Import as TQImport3XLS do
      begin
        SkipFirstRows := Self.XLSSkipRows;
        SkipFirstCols := Self.XLSSkipCols;
      end;

    TuneMap;
    Import.KeyColumns.Clear;
    for i := 0 to lvSelectedColumns.Items.Count - 1 do
      Import.KeyColumns.Add(lvSelectedColumns.Items[i].Caption);

    with Import.Formats do
    begin
      DecimalSeparator := Self.DecimalSeparator;
      ThousandSeparator := Self.ThousandSeparator;
      ShortDateFormat := Self.ShortDateFormat;
      LongDateFormat := Self.LongDateFormat;
      DateSeparator := Self.DateSeparator;
      ShortTimeFormat := Self.ShortTimeFormat;
      LongTimeFormat := Self.LongTimeFormat;
      TimeSeparator := Self.TimeSeparator;

      BooleanTrue.Assign(mmBooleanTrue.Lines);
      BooleanFalse.Assign(mmBooleanFalse.Lines);
      NullValues.Assign(mmNullValues.Lines);
    end;

    ApplyDataFormats(Import);

    Import.CommitAfterDone := CommitAfterDone;
    Import.CommitRecCount := CommitRecCount;
    Import.ImportRecCount := ImportRecCount * Integer(not chImportAllRecords.Checked);
    Import.ErrorLog := EnableErrorLog;

    fileDir := ExtractFileDir(ErrorLogFileName);
    if fileDir <> '' then
    begin
      if DirExists(fileDir) then
        Import.ErrorLogFileName := ErrorLogFileName
      else
        Import.ErrorLogFileName := ExtractFilePath(Application.ExeName) + ErrorLogFileName;
    end
    else
      Import.ErrorLogFileName := ExtractFilePath(Application.ExeName) + ErrorLogFileName;

    Import.RewriteErrorLogFile := RewriteErrorLogFile;
    Import.ShowErrorLog := ShowErrorLog;

    Import.ImportMode := ImportMode;
    Import.AddType := AddType;

    Import.Execute;

    if Assigned(FProgress) and not FCloseAfterImport then
      while FProgress.ModalResult <> mrOk do
        Application.ProcessMessages;
    if FCloseAfterImport then
      Close;
  finally
    if Assigned(FProgress) then
    begin
      FProgress.Free;
      FProgress := nil;
    end;
  end;
end;

procedure TQImport3WizardF.TuneFormats;
begin
  pgImport.ActivePage := tsFormats;
  pgFieldOptions.ActivePage := tsFieldTuning;
  if FImportType = aiAccess
    then laStep_06.Caption := Format(QImportLoadStr(QIW_Step), [3,4])
    else laStep_06.Caption := Format(QImportLoadStr(QIW_Step), [2,3]);
  pgFormats.ActivePage := tshBaseFormats;
  TuneButtons;
end;

procedure TQImport3WizardF.TuneMap;
var
  i: Integer;
begin
  case FImportType of
    aiXLS: begin
      impXLS.Map.Clear;
      for i := 0 to lvXLSFields.Items.Count - 1 do
        if TMapRow(lvXLSFields.Items[i].Data).AsString <> EmptyStr then
          impXLS.Map.Values[lvXLSFields.Items[i].Caption] :=
            TMapRow(lvXLSFields.Items[i].Data).AsString;
    end;
    aiDBF: begin
      impDBF.Map.Clear;
      for i := 0 to lstDBFMap.Items.Count - 1 do
        impDBF.Map.Values[lstDBFMap.Items[i].Caption] :=
          lstDBFMap.Items[i].SubItems[1];
    end;
    {$IFDEF XLSX}
    {$IFDEF VCL6}
    aiXlsx: begin
      impXlsx.Map.Clear;
      for i := 0 to lvXlsxFields.Items.Count - 1 do
        if lvXlsxFields.Items[i].SubItems[0] <> '' then
          impXlsx.Map.Values[lvXlsxFields.Items[i].Caption] :=
            lvXlsxFields.Items[i].SubItems[0];
    end;
    {$ENDIF}
    {$ENDIF}
    {$IFDEF DOCX}
    {$IFDEF VCL6}
    aiDocx: begin
      impDocx.Map.Clear;
      for i := 0 to lvDocxFields.Items.Count - 1 do
        if StrToIntDef(lvDocxFields.Items[i].SubItems[0], 0) > 0 then
          impDocx.Map.Values[lvDocxFields.Items[i].Caption] :=
            IntToStr(StrToInt(lvDocxFields.Items[i].SubItems[0]));
    end;
    {$ENDIF}
    {$ENDIF}
    {$IFDEF ODS}
    {$IFDEF VCL6}
    aiODS: begin
      impODS.Map.Clear;
      for i := 0 to lvODSFields.Items.Count - 1 do
        if lvODSFields.Items[i].SubItems[0] <> '' then
          impODS.Map.Values[lvODSFields.Items[i].Caption] :=
            lvODSFields.Items[i].SubItems[0];
    end;
    {$ENDIF}
    {$ENDIF}
    {$IFDEF ODT}
    {$IFDEF VCL6}
    aiODT: begin
      impODT.Map.Clear;
      for i := 0 to lvODTFields.Items.Count - 1 do
        if lvODTFields.Items[i].SubItems[0] <> '' then
          impODT.Map.Values[lvODTFields.Items[i].Caption] :=
            lvODTFields.Items[i].SubItems[0];
    end;
    {$ENDIF}
    {$ENDIF}
    {$IFDEF HTML}
    {$IFDEF VCL6}
    aiHTML: begin
      impHTML.Map.Clear;
      for i := 0 to lvHTMLFields.Items.Count - 1 do
        if StrToIntDef(lvHTMLFields.Items[i].SubItems[0], 0) > 0 then
          impHTML.Map.Values[lvHTMLFields.Items[i].Caption] :=
            IntToStr(StrToInt(lvHTMLFields.Items[i].SubItems[0]));
    end;
    {$ENDIF}
    {$ENDIF}
    aiXML: begin
      impXML.Map.Clear;
      for i := 0 to lvXMLMap.Items.Count - 1 do
        impXML.Map.Values[lvXMLMap.Items[i].Caption] :=
          lvXMLMap.Items[i].SubItems[1];
    end;
    {$IFDEF XMLDOC}
    {$IFDEF VCL6}
    aiXMLDoc: begin
      impXMLDoc.Map.Clear;
      for i := 0 to lvXMLDocFields.Items.Count - 1 do
        if StrToIntDef(lvXMLDocFields.Items[i].SubItems[0], 0) > 0 then
          impXMLDoc.Map.Values[lvXMLDocFields.Items[i].Caption] :=
            IntToStr(StrToInt(lvXMLDocFields.Items[i].SubItems[0]));
    end;
    {$ENDIF}
    {$ENDIF}
    aiTXT: begin
      impASCII.Map.Clear;
      for i := 0 to lvTXTFields.Items.Count - 1 do
        if StrToIntDef(lvTXTFields.Items[i].SubItems[1], 0) > 0 then
          impASCII.Map.Values[lvTXTFields.Items[i].Caption] :=
            Format('%s;%s', [lvTXTFields.Items[i].SubItems[0],
                             lvTXTFields.Items[i].SubItems[1]]);
    end;
    aiCSV: begin
      impASCII.Map.Clear;
      for i := 0 to lvCSVFields.Items.Count - 1 do
        if StrToIntDef(lvCSVFields.Items[i].SubItems[0], 0) > 0 then
          impASCII.Map.Values[lvCSVFields.Items[i].Caption] :=
            IntToStr(StrToInt(lvCSVFields.Items[i].SubItems[0]));
    end;
    {$IFDEF ADO}
    aiAccess: begin
      impAccess.Map.Clear;
      for i := 0 to lstAccessMap.Items.Count - 1 do
        impAccess.Map.Values[lstAccessMap.Items[i].Caption] :=
          lstAccessMap.Items[i].SubItems[1];
    end;
    {$ENDIF}
  end;
end;

procedure TQImport3WizardF.SetTitle;
begin
  if FileName <> EmptyStr then
    Caption := QImportLoadStr(QIW_ImportWizard) + ' - ' +
      Format(QImportLoadStr(QIW_ImportFrom) + ' %s', [ExtractFileName(FileName)])
  else Caption := QImportLoadStr(QIW_ImportWizard);
end;

procedure TQImport3WizardF.bHelpClick(Sender: TObject);
begin
  if FileExists(ExtractFilePath(Application.ExeName) + QI_WIZARD_HELP) then
  begin
    HelpFile := ExtractFilePath(Application.ExeName) + QI_WIZARD_HELP;
    Application.HelpCommand(HELP_CONTENTS, 0);
  end;
end;

procedure TQImport3WizardF.SetEnabledDataFormatControls;
var
  FGenStep: Integer;
begin
  // generator
  edtGeneratorValue.Enabled := (edtConstantValue.Text = EmptyStr) and
    (edtNullValue.Text = EmptyStr) and (edtDefaultValue.Text = EmptyStr) and
    (edtLeftQuote.Text = EmptyStr) and (edtRightQuote.Text = EmptyStr) and
    (TQuoteAction(cmbQuoteAction.ItemIndex) = qaNone) and
    (TQImportCharCase(cmbCharCase.ItemIndex) = iccNone) and
    (TQImportCharSet(cmbCharSet.ItemIndex) = icsNone) and
    (mScript.Lines.Text = EmptyStr);
  laGeneratorValue.Enabled := edtGeneratorValue.Enabled;
  edtGeneratorStep.Enabled := edtGeneratorValue.Enabled;
  laGeneratorStep.Enabled := edtGeneratorValue.Enabled;

  FGenStep := StrToIntDef(edtGeneratorStep.Text, 0);
  //constant
  edtConstantValue.Enabled := (FGenStep = 0) and
    (edtNullValue.Text = EmptyStr) and (edtDefaultValue.Text = EmptyStr) and
    (mScript.Lines.Text = EmptyStr);
  laConstantValue.Enabled := edtConstantValue.Enabled;
  // default
  edtNullValue.Enabled := (FGenStep = 0) and (edtConstantValue.Text = EmptyStr)
    and (mScript.Lines.Text = EmptyStr);
  laNullValue.Enabled := edtNullValue.Enabled;
  edtDefaultValue.Enabled := edtNullValue.Enabled;
  laDefaultValue.Enabled := edtNullValue.Enabled;
  // quotation
  edtLeftQuote.Enabled := (FGenStep = 0) and (mScript.Lines.Text = EmptyStr);
  laLeftQuote.Enabled := edtLeftQuote.Enabled;
  edtRightQuote.Enabled := edtLeftQuote.Enabled;
  laRightQuote.Enabled := edtLeftQuote.Enabled;
  cmbQuoteAction.Enabled := edtLeftQuote.Enabled;
  laQuoteAction.Enabled := edtLeftQuote.Enabled;
  // AnsiString coversion
  cmbCharCase.Enabled := (FGenStep = 0) and (mScript.Lines.Text = EmptyStr);
  laCharCase.Enabled := cmbCharCase.Enabled;
  cmbCharSet.Enabled := cmbCharCase.Enabled;
  laCharSet.Enabled := cmbCharCase.Enabled;
end;

procedure TQImport3WizardF.edtGeneratorValueChange(Sender: TObject);
begin
  if Assigned(lstFormatFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFormatFields.Selected.Data).GeneratorValue :=
      StrToIntDef(edtGeneratorValue.Text, 0);
    SetEnabledDataFormatControls;
  end;
end;

procedure TQImport3WizardF.edtGeneratorStepChange(Sender: TObject);
begin
  if Assigned(lstFormatFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFormatFields.Selected.Data).GeneratorStep :=
      StrToIntDef(edtGeneratorStep.Text, 0);
    SetEnabledDataFormatControls;
  end;
end;

procedure TQImport3WizardF.edtConstantValueChange(Sender: TObject);
begin
  if Assigned(lstFormatFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFormatFields.Selected.Data).ConstantValue :=
      edtConstantValue.Text;
    SetEnabledDataFormatControls;
  end;
end;

procedure TQImport3WizardF.edtNullValueChange(Sender: TObject);
begin
  if Assigned(lstFormatFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFormatFields.Selected.Data).NullValue :=
      edtNullValue.Text;
    SetEnabledDataFormatControls;
  end;
end;

procedure TQImport3WizardF.edtDefaultValueChange(Sender: TObject);
begin
  if Assigned(lstFormatFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFormatFields.Selected.Data).DefaultValue :=
      edtDefaultValue.Text;
    SetEnabledDataFormatControls;
  end;
end;

procedure TQImport3WizardF.edtLeftQuoteChange(Sender: TObject);
begin
  if Assigned(lstFormatFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFormatFields.Selected.Data).LeftQuote :=
      edtLeftQuote.Text;
    SetEnabledDataFormatControls;
  end;
end;

procedure TQImport3WizardF.edtRightQuoteChange(Sender: TObject);
begin
  if Assigned(lstFormatFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFormatFields.Selected.Data).RightQuote :=
      edtRightQuote.Text;
    SetEnabledDataFormatControls;
  end;
end;

procedure TQImport3WizardF.cmbQuoteActionChange(Sender: TObject);
begin
  if Assigned(lstFormatFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFormatFields.Selected.Data).QuoteAction :=
      TQuoteAction(cmbQuoteAction.ItemIndex);
    SetEnabledDataFormatControls;
  end;
end;

procedure TQImport3WizardF.cmbCharCaseChange(Sender: TObject);
begin
  if Assigned(lstFormatFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFormatFields.Selected.Data).CharCase :=
      TQImportCharCase(cmbCharCase.ItemIndex);
    SetEnabledDataFormatControls;
  end;
end;

procedure TQImport3WizardF.cmbCharSetChange(Sender: TObject);
begin
  if Assigned(lstFormatFields.Selected) and not FLoadingFormatItem then begin
    TQImportFieldFormat(lstFormatFields.Selected.Data).CharSet :=
      TQImportCharSet(cmbCharSet.ItemIndex);
    SetEnabledDataFormatControls;
  end;
end;

procedure TQImport3WizardF.ApplyDataFormats(AImport: TQImport3);
var
  i, j: Integer;
begin
  if lstFormatFields.Items.Count = 0 then Exit;
  for i := 0 to lstFormatFields.Items.Count - 1 do begin
    j := AImport.FieldFormats.IndexByName(lstFormatFields.Items[i].Caption);
    if j > -1 then
      AImport.FieldFormats[j].Assign(TQImportFieldFormat(lstFormatFields.Items[i].Data))
    else begin
      if not TQImportFieldFormat(lstFormatFields.Items[i].Data).IsDefaultValues then
        with AImport.FieldFormats.Add do
          Assign(TQImportFieldFormat(lstFormatFields.Items[i].Data));
    end;
  end;
end;

procedure TQImport3WizardF.lstFormatFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if (csDestroying in ComponentState) or (Change <> ctState) then Exit;

  if Assigned(Item) and Assigned(Item.Data) and
     Item.Selected and (Item <> FFormatItem) then ShowFormatItem(Item);
  TuneButtons;
end;

procedure TQImport3WizardF.pgFormatsChange(Sender: TObject);
begin
  if pgFormats.ActivePage = tshDataFormats then begin
    ActiveControl := lstFormatFields;
    if (lstFormatFields.Items.Count > 0) and
       ((not Assigned(lstFormatFields.Selected)) or (not lstFormatFields.Selected.Focused)) then begin
      lstFormatFields.Items[0].Focused := True;
      lstFormatFields.Items[0].Selected := True;
      ShowFormatItem(lstFormatFields.Items[0]);
    end;
  end;
end;

procedure TQImport3WizardF.LoadTemplateFromFile(const AFileName: qiString);
var
  Template: TIniFile;
  i, j, k, l: Integer;
  AStrings: TStrings;
  b: boolean;
  FF: TQImportFieldFormat;
  str: string;
  R: TQImportReplacement;
  textToFind, replaceWith: string;
  ignoreCase: Boolean;
  TxtPos,
  TxtSize: Integer;
begin
  if Wizard.ShowSaveLoadButtons then
  begin
    laTemplateFileName.Left := 168;
    laTemplateFileName.Width := 233;
  end
  else begin
    laTemplateFileNameTag.Left := 12;
    laTemplateFileName.Left := 12;
    laTemplateFileName.Width := 389;
  end;

  laTemplateFileNameTag.Visible :=
    FileExists(AFileName) and not Wizard.ShowSaveLoadButtons;

  if not FileExists(AFileName)
    then raise EQImportError.CreateFmt(QImportLoadStr(QIE_FileNotExists), [AFileName]);

  laTemplateFileName.Caption := MinimizeName(AFileName,
    laTemplateFileName.Canvas, laTemplateFileName.Width);

  FTmpFileName := '';
  Template := TIniFile.Create(AFileName);
  try
    AStrings := TStringList.Create;
    try
      with Template do
      begin
        // ImportType
        Self.FileName := ReadString(QIW_FIRST_STEP, QIW_FILE_NAME, EmptyStr);
        Wizard.FileName := Self.FileName;
        if not Wizard.GoToLastPage then
          GoToLastPage := ReadBool(QIW_FIRST_STEP, QIW_GO_TO_LAST_PAGE, False);
        if not Wizard.AutoSaveTemplate then
          AutoSaveTemplate := ReadBool(QIW_FIRST_STEP, QIW_AUTO_SAVE_TEMPLATE, False);
        ImportType := TAllowedImport(ReadInteger(QIW_FIRST_STEP, QIW_IMPORT_TYPE, 0));
        case ImportType of
          // ExcelOptions
          aiXLS: begin
            ReadSection(QIW_XLS_MAP, AStrings);
            for i := 0 to AStrings.Count - 1 do
            begin
              j := -1;
              for k := 0 to lvXLSFields.Items.Count - 1 do
              begin
                if AnsiCompareText(lvXLSFields.Items[k].Caption, AStrings[i]) = 0 then
                begin
                  j := k;
                  Break;
                end;
              end;
              if j > -1 then
                TMapRow(lvXLSFields.Items[j].Data).AsString :=
                  ReadString(QIW_XLS_MAP, AStrings[i], EmptyStr);
            end;
            XLSSkipCols := ReadInteger(QIW_XLS_OPTIONS, QIW_XLS_SKIP_COLS, 0);
            XLSSkipRows := ReadInteger(QIW_XLS_OPTIONS, QIW_XLS_SKIP_ROWS, 0);
          end;
          aiTXT:
            begin
              FTmpFileName := AFileName;
              TXTSkipLines :=ReadInteger(QIW_TXT_OPTIONS,
                            QIW_TXT_SKIP_LINES, 0);
              TXTEncoding := TQICharsetType( ReadInteger(QIW_TXT_OPTIONS,
                            QIW_TXT_ENCODING_OPTION, 0));

              ReadSection(QIW_TXT_MAP, AStrings);
              for i := 0 to AStrings.Count - 1 do
              begin
                j := -1;
                for k := 0 to lvTXTFields.Items.Count - 1 do
                begin
                  if AnsiCompareText(lvTXTFields.Items[k].Caption, AStrings[i]) = 0 then
                  begin
                    j := k;
                    Break;
                  end;
                end;
                if j > -1 then
                begin
                  TXTExtractPosSize(ReadString(QIW_TXT_MAP, AStrings[i],
                    EmptyStr), TxtPos, TxtSize);
                  lvTXTFields.Items[j].SubItems[0] := IntToStr(TxtPos);
                  lvTXTFields.Items[j].SubItems[1] := IntToStr(TxtSize);
                end;
              end;

            end;
          aiCSV: begin
            Comma := Str2Char(ReadString(QIW_CSV_OPTIONS, QIW_CSV_DELIMITER,
              Char2Str(GetListSeparator)), Comma);
            Quote := Str2Char(ReadString(QIW_CSV_OPTIONS, QIW_CSV_QUOTE, '"'),
              Quote);
            CSVSkipLines := ReadInteger(QIW_CSV_OPTIONS, QIW_CSV_SKIP_LINES, 0);

            ReadSection(QIW_CSV_MAP, AStrings);
            for i := 0 to AStrings.Count - 1 do
            begin
              j := -1;
              for k := 0 to lvCSVFields.Items.Count - 1 do
              begin
                if AnsiCompareText(lvCSVFields.Items[k].Caption, AStrings[i]) = 0 then
                begin
                  j := k;
                  Break;
                end;
              end;
              if j > -1 then
              begin
                str := ReadString(QIW_CSV_MAP, AStrings[i], EmptyStr);
                if StrToIntDef(str, 0) > 0 then
                  lvCSVFields.Items[j].SubItems[0] := str;
              end;
            end;
          end;
          aiDBF: begin
            DBFSkipDeleted := ReadBool(QIW_DBF_OPTIONS, QIW_DBF_SKIP_DELETED, True);
            lstDBFMap.Items.Clear;
            ReadSection(QIW_DBF_MAP, AStrings);
            for i := 0 to AStrings.Count - 1 do
            begin
              b := False;
              for j := 0 to QImportDestinationColCount(False,
                              ImportDestination, DataSet, DBGrid, ListView,
                              StringGrid) - 1 do
                b := b or (AnsiCompareText(AStrings[i],
                  QImportDestinationColName(False, ImportDestination, DataSet,
                    DBGrid, ListView, StringGrid, GridCaptionRow, j)) = 0);
              if not b then Continue;
              with lstDBFMap.Items.Add do
              begin
                Caption := AStrings[i];
                SubItems.Add('=');
                SubItems.Add(ReadString(QIW_DBF_MAP, AStrings[i], EmptyStr));
              end;
              for j := 0 to lstDBFDataSet.Items.Count - 1 do
              begin
                if AnsiCompareText(lstDBFDataSet.Items[j].Caption, AStrings[i]) = 0 then
                begin
                  lstDBFDataSet.Items[j].Delete;
                  Break;
                end;
              end;
            end;
            if lstDBFMap.Items.Count > 0 then
            begin
              lstDBFMap.Items[0].Focused := True;
              lstDBFMap.Items[0].Selected := True;
            end;
          end;
          {!TODO}
          {$IFDEF HTML}
          {$IFDEF VCL6}
          aiHTML: begin
            HTMLSkipLines := ReadInteger(QIW_HTML_OPTIONS, QIW_HTML_SKIP_LINES, 0);
            ReadSection(QIW_HTML_MAP, AStrings);
            for i := 0 to AStrings.Count - 1 do
            begin
              j := -1;
              for k := 0 to lvHTMLFields.Items.Count - 1 do
              begin
                if AnsiCompareText(lvHTMLFields.Items[k].Caption, AStrings[i]) = 0 then
                begin
                  j := k;
                  Break;
                end;
              end;
              if j > -1 then
              begin
                str := ReadString(QIW_HTML_MAP, AStrings[i], EmptyStr);
                if StrToIntDef(str, 0) > 0 then
                  lvHTMLFields.Items[j].SubItems[0] := str;
              end;
            end;
          end;
          {$ENDIF}
          {$ENDIF}
          {$IFDEF ODS}
          {$IFDEF VCL6}
          aiODS: begin
            ODSSkipLines := ReadInteger(QIW_ODS_OPTIONS, QIW_ODS_SKIP_LINES, 0);
            ODSSheetName := AnsiString(ReadString(QIW_ODS_OPTIONS, QIW_ODS_SHEET_NAME, 'Sheet1'));
            ReadSection(QIW_ODS_MAP, AStrings);
            for i := 0 to AStrings.Count - 1 do
            begin
              j := -1;
              for k := 0 to lvODSFields.Items.Count - 1 do
              begin
                if AnsiCompareText(lvODSFields.Items[k].Caption, AStrings[i]) = 0 then
                begin
                  j := k;
                  Break;
                end;
              end;
              if j > -1 then
              begin
                str := ReadString(QIW_ODS_MAP, AStrings[i], EmptyStr);
                if str <> EmptyStr then
                  lvODSFields.Items[j].SubItems[0] := str;
              end;
            end;
          end;
          {$ENDIF}
          {$ENDIF}
          {$IFDEF ODT}
          {$IFDEF VCL6}
          aiODT: begin
            ODTSkipLines := ReadInteger(QIW_ODT_OPTIONS, QIW_ODT_SKIP_LINES, 0);
            ODTSheetName := AnsiString(ReadString(QIW_ODT_OPTIONS, QIW_ODT_SHEET_NAME, 'Table1'));
            ODTUseHeader := ReadBool(QIW_ODT_OPTIONS, QIW_ODT_USE_HEADER, false);
            cbODTUseHeader.Checked := ODTUseHeader;
            ReadSection(QIW_ODT_MAP, AStrings);
            for i := 0 to AStrings.Count - 1 do
            begin
              j := -1;
              for k := 0 to lvODTFields.Items.Count - 1 do
              begin
                if AnsiCompareText(lvODTFields.Items[k].Caption, AStrings[i]) = 0 then
                begin
                  j := k;
                  Break;
                end;
              end;
              if j > -1 then
              begin
                str := ReadString(QIW_ODT_MAP, AStrings[i], EmptyStr);
                if str <> EmptyStr then
                  lvODTFields.Items[j].SubItems[0] := str;
              end;
            end;
          end;
          {$ENDIF}
          {$ENDIF}
          {$IFDEF XLSX}
          {$IFDEF VCL6}
          aiXlsx: begin
            XlsxSkipLines := ReadInteger(QIW_XLSX_OPTIONS, QIW_XLSX_SKIP_LINES, 0);
            XlsxSheetName := ReadString(QIW_XLSX_OPTIONS, QIW_XLSX_SHEET_NAME, 'Sheet1');
            XlsxNeedFillMerge := ReadBool(QIW_XLSX_OPTIONS, QIW_XLSX_NEED_FILLMERGE, false);
            XlsxLoadHiddenSheet := ReadBool(QIW_XLSX_OPTIONS, QIW_XLSX_LOAD_HIDDENSHEET, false);
            ReadSection(QIW_XLSX_MAP, AStrings);
            for i := 0 to AStrings.Count - 1 do
            begin
              j := -1;
              for k := 0 to lvXlsxFields.Items.Count - 1 do
              begin
                if AnsiCompareText(lvXLSXFields.Items[k].Caption, AStrings[i]) = 0 then
                begin
                  j := k;
                  Break;
                end;
              end;
              if j > -1 then
              begin
                str := ReadString(QIW_XLSX_MAP, AStrings[i], EmptyStr);
                if str <> EmptyStr then
                  lvXlsxFields.Items[j].SubItems[0] := str;
              end;
            end;
          end;
          {$ENDIF}
          {$ENDIF}
          {$IFDEF DOCX}
          {$IFDEF VCL6}
          aiDocx: begin
            DocxSkipLines := ReadInteger(QIW_DOCX_OPTIONS, QIW_DOCX_SKIP_LINES, 0);
            DocxTableNumber := ReadInteger(QIW_DOCX_OPTIONS, QIW_DOCX_TABLE_NUMBER, 0);
            ReadSection(QIW_DOCX_MAP, AStrings);
            for i := 0 to AStrings.Count - 1 do
            begin
              j := -1;
              for k := 0 to lvDocxFields.Items.Count - 1 do
              begin
                if AnsiCompareText(lvDocxFields.Items[k].Caption, AStrings[i]) = 0 then
                begin
                  j := k;
                  Break;
                end;
              end;
              if j > -1 then
              begin
                str := ReadString(QIW_DOCX_MAP, AStrings[i], EmptyStr);
                if str <> EmptyStr then
                  lvDocxFields.Items[j].SubItems[0] := str;
              end;
            end;
          end;
          {$ENDIF}
          {$ENDIF}
          aiXML: begin 
            XMLWriteOnFly := ReadBool(QIW_XML_OPTIONS, QIW_XML_WRITE_ON_FLY, False);
            lvXMLMap.Items.Clear;
            ReadSection(QIW_XML_MAP, AStrings);
            for i := 0 to AStrings.Count - 1 do
            begin
              b := False;
              for j := 0 to QImportDestinationColCount(False,
                              ImportDestination, DataSet, DBGrid, ListView,
                              StringGrid) - 1 do
                b := b or (AnsiCompareText(AStrings[i],
                  QImportDestinationColName(False, ImportDestination, DataSet,
                    DBGrid, ListView, StringGrid, GridCaptionRow, j)) = 0);
              if not b then Continue;
              with lvXMLMap.Items.Add do
              begin
                Caption := AStrings[i];
                SubItems.Add('=');
                SubItems.Add(ReadString(QIW_XML_MAP, AStrings[i], EmptyStr));
              end;
              for j := 0 to lvXMLDataSet.Items.Count - 1 do
              begin
                if AnsiCompareText(lvXMLDataSet.Items[j].Caption, AStrings[i]) = 0 then
                begin
                  lvXMLDataSet.Items[j].Delete;
                  Break;
                end;
              end;
            end;
            if lvXMLMap.Items.Count > 0 then
            begin
              lvXMLMap.Items[0].Focused := True;
              lvXMLMap.Items[0].Selected := True;
            end;
          end;
          {$IFDEF ADO}
          aiAccess: begin
            if AFileName <> EmptyStr then
            begin
              if Self.FileName <> EmptyStr then
              begin
                AccessPassword := ReadString(QIW_MDB_OPTIONS, QIW_MDB_PASSWORD, EmptyStr);
                AccessGetTableNames;
                AccessSourceType := TQImportAccessSourceType(ReadInteger(
                  QIW_MDB_OPTIONS, QIW_MDB_SOURCETYPE, 0));
                if AccessSourceType = isTable then
                begin
                  str := ReadString(QIW_MDB_OPTIONS, QIW_MDB_TABLENAME, EmptyStr);
                  i := lbAccessTables.Items.IndexOf(str);
                  if i < 0 then i := 0;
                  lbAccessTables.ItemIndex := i;
                  impAccess.SourceType := isTable;
                  impAccess.TableName := lbAccessTables.Items[lbAccessTables.ItemIndex];
                end
                else begin
                  memAccessSQL.Lines.Clear;
                  ReadSection(QIW_MDB_QUERY, AStrings);
                  for i := 0 to AStrings.Count - 1 do
                    memAccessSQL.Lines.Add(ReadString(QIW_MDB_QUERY, AStrings[i],
                      EmptyStr));
                  impAccess.SourceType := isSQL;
                  impAccess.SQL.Assign(memAccessSQL.Lines);
                end;
                AccessGetFieldNames;
                lstAccessMap.Items.Clear;
                ReadSection(QIW_MDB_MAP, AStrings);
                for i := 0 to AStrings.Count - 1 do
                begin
                  b := false;
                  for j := 0 to QImportDestinationColCount(false,
                                  ImportDestination, DataSet, DBGrid, ListView,
                                  StringGrid) - 1 do
                    b := b or (AnsiCompareText(AStrings[i],
                      QImportDestinationColName(false, ImportDestination,
                        DataSet, DBGrid, ListView, StringGrid,
                        GridCaptionRow, j)) = 0);
                  if not b then Continue;
                  with lstAccessMap.Items.Add do
                  begin
                    Caption := AStrings[i];
                    SubItems.Add('=');
                    SubItems.Add(ReadString(QIW_MDB_MAP, AStrings[i], EmptyStr));
                  end;
                  for j := 0 to lstAccessDataSet.Items.Count - 1 do
                    if AnsiCompareText(lstAccessDataSet.Items[j].Caption,
                                       AStrings[i]) = 0 then
                    begin
                      lstAccessDataSet.Items[j].Delete;
                      Break;
                    end;
                end;
                if lstAccessMap.Items.Count > 0 then
                begin
                  lstAccessMap.Items[0].Focused := true;
                  lstAccessMap.Items[0].Selected := true;
                end;
              end;
            end;
          end;
        {$ENDIF}
        {$IFDEF XMLDOC}
        {$IFDEF VCL6}
        aiXMLDoc:
        begin
          XMLDocXPath := ReadString(QIW_XML_DOC_OPTIONS, QIW_XML_DOC_XPATH,
            XMLDocXPath);
          XMLDocDataLocation := TXMLDataLocation(ReadInteger(QIW_XML_DOC_OPTIONS,
            QIW_XML_DOC_DATALOCATION, Integer(XMLDocDataLocation)));
          XMLDocSkipLines := ReadInteger(QIW_XML_DOC_OPTIONS,
            QIW_XML_DOC_SKIPLINES, XMLDocSkipLines);
          for I := 0 to lvXMLDocFields.Items.Count - 1 do
            lvXMLDocFields.Items[I].SubItems[0] := ReadString(QIW_XML_DOC_MAP,
              QIW_XML_DOC_MAP_LINE + IntToStr(I),
              lvXMLDocFields.Items[I].SubItems[0]);
        end;
        {$ENDIF}
        {$ENDIF}
        end;
        // Formats
        DecimalSeparator := Str2Char(ReadString(QIW_BASE_FORMATS,
          QIW_BF_DECIMAL_SEPARATOR, Char2Str(Wizard.Formats.DecimalSeparator)),
          Wizard.Formats.DecimalSeparator);
        ThousandSeparator := Str2Char(ReadString(QIW_BASE_FORMATS,
          QIW_BF_THOUSAND_SEPARATOR, Char2Str(Wizard.Formats.ThousandSeparator)),
          Wizard.Formats.ThousandSeparator);
        ShortDateFormat := ReadString(QIW_BASE_FORMATS, QIW_BF_SHORT_DATE_FORMAT,
          Wizard.Formats.ShortDateFormat);
        LongDateFormat := ReadString(QIW_BASE_FORMATS, QIW_BF_LONG_DATE_FORMAT,
          Wizard.Formats.LongDateFormat);
        DateSeparator := Str2Char(ReadString(QIW_BASE_FORMATS,
          QIW_BF_DATE_SEPARATOR, Char2Str(Wizard.Formats.DateSeparator)),
          Wizard.Formats.DateSeparator);
        ShortTimeFormat := ReadString(QIW_BASE_FORMATS, QIW_BF_SHORT_TIME_FORMAT,
          Wizard.Formats.ShortTimeFormat);
        LongTimeFormat := ReadString(QIW_BASE_FORMATS, QIW_BF_LONG_TIME_FORMAT,
          Wizard.Formats.LongTimeFormat);
        TimeSeparator := Str2Char(ReadString(QIW_BASE_FORMATS,
          QIW_BF_TIME_SEPARATOR, Char2Str(Wizard.Formats.TimeSeparator)),
          Wizard.Formats.TimeSeparator);
        ReadSection(QIW_BOOLEAN_TRUE, mmBooleanTrue.Lines);
        ReadSection(QIW_BOOLEAN_FALSE, mmBooleanFalse.Lines);
        ReadSection(QIW_NULL_VALUES, mmNullValues.Lines);
        // DataFormats
        ReadSections(AStrings);
        for i := 0 to AStrings.Count - 1 do
        begin
          if Pos(QIW_DATA_FORMATS, AStrings[i]) > 0 then
          begin
            for j := 0 to lstFormatFields.Items.Count - 1 do
              if AnsiCompareText(Copy(AStrings[i], Length(QIW_DATA_FORMATS) + 1,
                Length(AStrings[i])), lstFormatFields.Items[j].Caption) = 0 then
              begin
                FF := TQImportFieldFormat(lstFormatFields.Items[j].Data);
                FF.Replacements.Clear;
                FF.GeneratorValue := ReadInteger(AStrings[i],
                  QIW_DF_GENERATOR_VALUE, FF.GeneratorValue);
                FF.GeneratorStep := ReadInteger(AStrings[i],
                  QIW_DF_GENERATOR_STEP, FF.GeneratorStep);
                FF.ConstantValue := ReadString(AStrings[i],
                  QIW_DF_CONSTANT_VALUE, FF.ConstantValue);
                FF.NullValue := ReadString(AStrings[i], QIW_DF_NULL_VALUE,
                  FF.NullValue);
                FF.DefaultValue := ReadString(AStrings[i], QIW_DF_DEFAULT_VALUE,
                  FF.DefaultValue);
                FF.LeftQuote := ReadString(AStrings[i], QIW_DF_LEFT_QUOTE,
                  FF.LeftQuote);
                FF.RightQuote := ReadString(AStrings[i], QIW_DF_RIGHT_QUOTE,
                  FF.RightQuote);
                FF.QuoteAction := TQuoteAction(ReadInteger(AStrings[i],
                  QIW_DF_QUOTE_ACTION, Integer(FF.QuoteAction)));
                FF.CharCase := TQImportCharCase(ReadInteger(AStrings[i],
                  QIW_DF_CHAR_CASE, Integer(FF.CharCase)));
                FF.CharSet := TQImportCharSet(ReadInteger(AStrings[i],
                  QIW_DF_CHAR_SET, Integer(FF.CharSet)));
              end;
          end
          // Replacements
          else if Pos(QIW_REPLACEMENTS, AStrings[i]) = 1 then
          begin
            l := Length(AStrings[i]);
            while QImport3Common.CharInSet(AStrings[i][l], ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']) do Dec(l);
            str := Copy(AStrings[i], Length(QIW_REPLACEMENTS) + 1,
                   l - Length(QIW_REPLACEMENTS) - 1);

            for j := 0 to lstFormatFields.Items.Count - 1 do
              if AnsiCompareText(str, lstFormatFields.Items[j].Caption) = 0 then
              begin
                FF := TQImportFieldFormat(lstFormatFields.Items[j].Data);
                textToFind := ReadString(AStrings[i], QIW_RP_TEXT_TO_FIND, EmptyStr);
                replaceWith := ReadString(AStrings[i], QIW_RP_REPLACE_WITH, EmptyStr);
                ignoreCase := ReadBool(AStrings[i], QIW_RP_IGNORE_CASE, False);
                if not FF.Replacements.ItemExists(textToFind, replaceWith, ignoreCase) then
                begin
                  R := FF.Replacements.Add;
                  R.TextToFind := textToFind;
                  R.ReplaceWith := replaceWith;
                  R.IgnoreCase := ignoreCase;
                end;  
              end;
              ShowFormatItem(lstFormatFields.ItemFocused);
          end;
        end;
        // LastStep
        if not Wizard.CommitAfterDone then
          CommitAfterDone := ReadBool(QIW_LAST_STEP, QIW_COMMIT_AFTER_DONE,
            Wizard.CommitAfterDone);
        CommitRecCount := ReadInteger(QIW_LAST_STEP, QIW_COMMIT_REC_COUNT,
          Wizard.CommitRecCount);
        ImportRecCount := ReadInteger(QIW_LAST_STEP, QIW_IMPORT_REC_COUNT,
          Wizard.ImportRecCount);
        if not Wizard.CloseAfterImport then
          CloseAfterImport := ReadBool(QIW_LAST_STEP, QIW_CLOSE_AFTER_IMPORT,
            Wizard.CloseAfterImport);
        EnableErrorLog := ReadBool(QIW_LAST_STEP, QIW_ENABLE_ERROR_LOG,
          Wizard.ErrorLog);
        ErrorLogFileName := ReadString(QIW_LAST_STEP, QIW_ERROR_LOG_FILE_NAME,
          Wizard.ErrorLogFileName);
        if not Wizard.RewriteErrorLogFile then
          RewriteErrorLogFile := ReadBool(QIW_LAST_STEP, QIW_REWRITE_ERROR_LOG_FILE,
            Wizard.RewriteErrorLogFile);
        if not Wizard.ShowErrorLog then
          ShowErrorLog := ReadBool(QIW_LAST_STEP, QIW_SHOW_ERROR_LOG, True);

        AutoTrimValue := ReadBool(QIW_LAST_STEP, QIW_AUTO_TRIM_VALUE,
          Wizard.AutoTrimValue);

        ImportEmptyRows := ReadBool(QIW_LAST_STEP, QIW_IMPORT_EMPTY_ROWS,
          Wizard.ImportEmptyRows);

        ImportMode := TQImportMode(ReadInteger(QIW_LAST_STEP, QIW_IMPORT_MODE,
          Integer(Wizard.ImportMode)));
        AddType := TQImportAddType(ReadInteger(QIW_LAST_STEP, QIW_ADD_TYPE,
          Integer(Wizard.AddType)));
        AStrings.Clear;
        ReadSection(QIW_KEY_COLUMNS, AStrings);
        FillKeyColumns(AStrings);
      end;
    finally
      AStrings.Free;
    end;
  finally
    Template.Free
  end;

  if Wizard.GoToLastPage then
  begin
    if (ImportType = aiTXT) and not TXTReady then
      TXTTune;
    if (((ImportType = aiXLS) and XLSReady) or
     ((ImportType = aiDBF) and DBFReady) or
     {$IFDEF XLSX}
     {$IFDEF VCL6}
     ((ImportType = aiXlsx) and XlsxReady) or
     {$ENDIF}{$ENDIF}

     {$IFDEF DOCX}
     {$IFDEF VCL6}
     ((ImportType = aiDocx) and DocxReady) or
     {$ENDIF}{$ENDIF}

     {$IFDEF ODS}
     {$IFDEF VCL6}
     ((ImportType = aiODS) and ODSReady) or
     {$ENDIF}{$ENDIF}

     {$IFDEF ODT}
     {$IFDEF VCL6}
     ((ImportType = aiODT) and ODTReady) or
     {$ENDIF}{$ENDIF}

     {$IFDEF HTML}
     {$IFDEF VCL6}
     ((ImportType = aiHTML) and HTMLReady) or
     {$ENDIF}{$ENDIF}

     ((ImportType = aiXML) and XMLReady) or
     {$IFDEF XMLDOC}
     {$IFDEF VCL6}
     ((ImportType = aiXMLDoc) and XMLDocReady) or
     {$ENDIF}{$ENDIF}

     ((ImportType = aiTXT) and TXTReady) or
     ((ImportType = aiCSV) and CSVReady)) then
    Step := 3
    {$IFDEF ADO}
    else if ((ImportType = aiAccess) and AccessReady) then
      Step := 4;
    {$ENDIF}
  end;
end;

procedure TQImport3WizardF.SaveTemplateToFile(const AFileName: qiString);
var
  Template: TIniFile;
  i, j: Integer;
  str: string;
  ff: TQImportFieldFormat;
begin
  if Trim(AFileName) = EmptyStr then Exit;
  Template := TIniFile.Create(AFileName);
  try
    with Template do begin
      ClearIniFile(Template);
      WriteInteger(QIW_FIRST_STEP, QIW_IMPORT_TYPE, Integer(ImportType));
      WriteString(QIW_FIRST_STEP, QIW_FILE_NAME, Self.FileName);
      WriteBool(QIW_FIRST_STEP, QIW_GO_TO_LAST_PAGE, GoToLastPage);
      WriteBool(QIW_FIRST_STEP, QIW_AUTO_SAVE_TEMPLATE, AutoSaveTemplate);
      case ImportType of
        aiXLS: begin
          for i := 0 to lvXLSFields.Items.Count - 1 do
            if TMapRow(lvXLSFields.Items[i].Data).AsString <> EmptyStr
              then WriteString(QIW_XLS_MAP, lvXLSFields.Items[i].Caption,
                TMapRow(lvXLSFields.Items[i].Data).AsString);
          WriteInteger(QIW_XLS_OPTIONS, QIW_XLS_SKIP_COLS, XLSSkipCols);
          WriteInteger(QIW_XLS_OPTIONS, QIW_XLS_SKIP_ROWS, XLSSkipRows);
        end;
        aiTXT: begin
          for i := 0 to lvTXTFields.Items.Count - 1 do
            if StrToIntDef(lvTXTFields.Items[i].SubItems[1], 0) > 0 then
              WriteString(QIW_TXT_MAP, lvTXTFields.Items[i].Caption,
                Format('%s;%s', [lvTXTFields.Items[i].SubItems[0],
                                 lvTXTFields.Items[i].SubItems[1]]));
          WriteInteger(QIW_TXT_OPTIONS, QIW_TXT_SKIP_LINES,
            TXTSkipLines);
          WriteInteger(QIW_TXT_OPTIONS, QIW_TXT_ENCODING_OPTION,
            Integer(TXTEncoding) );
        end;
        aiCSV: begin
          for i := 0 to lvCSVFields.Items.Count - 1 do
            if StrToIntDef(lvCSVFields.Items[i].SubItems[0], 0) > 0 then
              WriteString(QIW_CSV_MAP, lvCSVFields.Items[i].Caption,
                lvCSVFields.Items[i].SubItems[0]);
          WriteInteger(QIW_CSV_OPTIONS, QIW_CSV_SKIP_LINES, CSVSkipLines);
          WriteString(QIW_CSV_OPTIONS, QIW_CSV_DELIMITER, Char2Str(Comma));
          WriteString(QIW_CSV_OPTIONS, QIW_CSV_QUOTE, Char2Str(Quote));
        end;
        aiDBF: begin
          WriteBool(QIW_DBF_OPTIONS, QIW_DBF_SKIP_DELETED, DBFSkipDeleted);
          for i := 0 to lstDBFMap.Items.Count - 1 do
            WriteString(QIW_DBF_MAP, lstDBFMap.Items[i].Caption,
              lstDBFMap.Items[i].SubItems[1]);
        end;
        {!TODO}
        {$IFDEF HTML}
        {$IFDEF VCL6}
        aiHTML: begin
          for i := 0 to lvHTMLFields.Items.Count - 1 do
            if StrToIntDef(lvHTMLFields.Items[i].SubItems[0], 0) > 0 then
              WriteString(QIW_HTML_MAP, lvHTMLFields.Items[i].Caption,
                lvHTMLFields.Items[i].SubItems[0]);
          WriteInteger(QIW_HTML_OPTIONS, QIW_HTML_SKIP_LINES, HTMLSkipLines);
        end;
        {$ENDIF}
        {$ENDIF}
        {$IFDEF ODS}
        {$IFDEF VCL6}
        aiODS: begin
          for i := 0 to lvODSFields.Items.Count - 1 do
            if lvODSFields.Items[i].SubItems[0] <> EmptyStr then
              WriteString(QIW_ODS_MAP, lvODSFields.Items[i].Caption,
                lvODSFields.Items[i].SubItems[0]);
          WriteInteger(QIW_ODS_OPTIONS, QIW_ODS_SKIP_LINES, ODSSkipLines);
          WriteString(QIW_ODS_OPTIONS, QIW_ODS_SHEET_NAME, string(ODSSheetName));
        end;
        {$ENDIF}
        {$ENDIF}
        {$IFDEF ODT}
        {$IFDEF VCL6}
        aiODT: begin
          for i := 0 to lvODTFields.Items.Count - 1 do
            if lvODTFields.Items[i].SubItems[0] <> EmptyStr then
              WriteString(QIW_ODT_MAP, lvODTFields.Items[i].Caption,
                lvODTFields.Items[i].SubItems[0]);
          WriteInteger(QIW_ODT_OPTIONS, QIW_ODT_SKIP_LINES, ODTSkipLines);
          WriteString(QIW_ODT_OPTIONS, QIW_ODT_SHEET_NAME, string(ODTSheetName));
          WriteBool(QIW_ODT_OPTIONS, QIW_ODT_USE_HEADER, ODTUseHeader);
        end;
        {$ENDIF}
        {$ENDIF}
        {$IFDEF XLSX}
        {$IFDEF VCL6}
        aiXlsx: begin
          for i := 0 to lvXlsxFields.Items.Count - 1 do
            if lvXlsxFields.Items[i].SubItems[0] <> EmptyStr then
              WriteString(QIW_XLSX_MAP, lvXlsxFields.Items[i].Caption,
                lvXlsxFields.Items[i].SubItems[0]);
          WriteInteger(QIW_XLSX_OPTIONS, QIW_XLSX_SKIP_LINES, XlsxSkipLines);
          WriteString(QIW_XLSX_OPTIONS, QIW_XLSX_SHEET_NAME, XlsxSheetName);
          WriteBool(QIW_XLSX_OPTIONS, QIW_XLSX_NEED_FILLMERGE, XlsxNeedFillMerge);
          WriteBool(QIW_XLSX_OPTIONS, QIW_XLSX_LOAD_HIDDENSHEET, XlsxLoadHiddenSheet);
        end;
        {$ENDIF}
        {$ENDIF}
        {$IFDEF DOCX}
        {$IFDEF VCL6}
        aiDocx: begin
          for i := 0 to lvDocxFields.Items.Count - 1 do
            if lvDocxFields.Items[i].SubItems[0] <> EmptyStr then
              WriteString(QIW_DOCX_MAP, lvDocxFields.Items[i].Caption,
                lvDocxFields.Items[i].SubItems[0]);
          WriteInteger(QIW_DOCX_OPTIONS, QIW_DOCX_SKIP_LINES, DocxSkipLines);
          WriteInteger(QIW_DOCX_OPTIONS, QIW_DOCX_TABLE_NUMBER, DocxTableNumber);
        end;
        {$ENDIF}
        {$ENDIF}
        aiXML: begin
          WriteBool(QIW_XML_OPTIONS, QIW_XML_WRITE_ON_FLY, XMLWriteOnFly);
          for i := 0 to lvXMLMap.Items.Count - 1 do
            WriteString(QIW_XML_MAP, lvXMLMap.Items[i].Caption,
              lvXMLMap.Items[i].SubItems[1]);
        end;
        {$IFDEF ADO}
        aiAccess: begin
          WriteString(QIW_MDB_OPTIONS, QIW_MDB_PASSWORD, AccessPassword);
          WriteInteger(QIW_MDB_OPTIONS, QIW_MDB_SOURCETYPE,
            Integer(AccessSourceType));
          if lbAccessTables.ItemIndex >= 0 then
          begin
            WriteString(QIW_MDB_OPTIONS, QIW_MDB_TABLENAME,
              lbAccessTables.Items[lbAccessTables.ItemIndex]);
          end;    
          for i := 0 to memAccessSQL.Lines.Count - 1 do
            WriteString(QIW_MDB_QUERY, QIW_MDB_SQL_LINE + IntToStr(i),
              memAccessSQL.Lines[i]);
          for i := 0 to lstAccessMap.Items.Count - 1 do
            WriteString(QIW_MDB_MAP, lstAccessMap.Items[i].Caption,
              lstAccessMap.Items[i].SubItems[1]);
        end;
        {$ENDIF}
        {$IFDEF XMLDOC}
        {$IFDEF VCL6}
        aiXMLDoc:
        begin
          WriteString(QIW_XML_DOC_OPTIONS, QIW_XML_DOC_XPATH, XMLDocXPath);
          WriteInteger(QIW_XML_DOC_OPTIONS, QIW_XML_DOC_DATALOCATION,
            Integer(XMLDocDataLocation));
          WriteInteger(QIW_XML_DOC_OPTIONS, QIW_XML_DOC_SKIPLINES,
            XMLDocSkipLines);
          for I := 0 to lvXMLDocFields.Items.Count - 1 do
            WriteString(QIW_XML_DOC_MAP, QIW_XML_DOC_MAP_LINE + IntToStr(I),
              lvXMLDocFields.Items[I].SubItems[0]);
        end;
        {$ENDIF}
        {$ENDIF}
      end;
      // Formats
      WriteString(QIW_BASE_FORMATS, QIW_BF_DECIMAL_SEPARATOR,
        Char2Str(Self.DecimalSeparator));
      WriteString(QIW_BASE_FORMATS, QIW_BF_THOUSAND_SEPARATOR,
        Char2Str(Self.ThousandSeparator));
      WriteString(QIW_BASE_FORMATS, QIW_BF_SHORT_DATE_FORMAT,
        Self.ShortDateFormat);
      WriteString(QIW_BASE_FORMATS, QIW_BF_LONG_DATE_FORMAT,
        Self.LongDateFormat);
      WriteString(QIW_BASE_FORMATS, QIW_BF_DATE_SEPARATOR,
        Char2Str(Self.DateSeparator));
      WriteString(QIW_BASE_FORMATS, QIW_BF_SHORT_TIME_FORMAT,
        Self.ShortTimeFormat);
      WriteString(QIW_BASE_FORMATS, QIW_BF_LONG_TIME_FORMAT,
        Self.LongTimeFormat);
      WriteString(QIW_BASE_FORMATS, QIW_BF_TIME_SEPARATOR,
        Char2Str(Self.TimeSeparator));
      for i := 0 to mmBooleanTrue.Lines.Count - 1 do
        WriteString(QIW_BOOLEAN_TRUE, mmBooleanTrue.Lines[i], EmptyStr);
      for i := 0 to mmBooleanFalse.Lines.Count - 1 do
        WriteString(QIW_BOOLEAN_FALSE, mmBooleanFalse.Lines[i], EmptyStr);
      for i := 0 to mmNullValues.Lines.Count - 1 do
        WriteString(QIW_NULL_VALUES, mmNullValues.Lines[i], EmptyStr);
      for i := 0 to lstFormatFields.Items.Count - 1 do begin
        ff := TQImportFieldFormat(lstFormatFields.Items[i].Data);
        str := QIW_DATA_FORMATS + AnsiUpperCase(ff.FieldName);
        if ff.GeneratorStep <> 0 then begin
          WriteInteger(str, QIW_DF_GENERATOR_VALUE, ff.GeneratorValue);
          WriteInteger(str, QIW_DF_GENERATOR_STEP, ff.GeneratorStep);
          Continue;
        end;
        if ff.ConstantValue <> EmptyStr then
          WriteString(str, QIW_DF_CONSTANT_VALUE, ff.ConstantValue);
        if (ff.ConstantValue = EmptyStr) and
           (ff.NullValue <> EmptyStr) and (ff.DefaultValue <> EmptyStr) then
        begin
          WriteString(str, QIW_DF_NULL_VALUE, ff.NullValue);
          WriteString(str, QIW_DF_DEFAULT_VALUE, ff.DefaultValue);
        end;
        if ff.LeftQuote <> EmptyStr then
          WriteString(str, QIW_DF_LEFT_QUOTE, ff.LeftQuote);
        if ff.RightQuote <> EmptyStr then
          WriteString(str, QIW_DF_RIGHT_QUOTE, ff.RightQuote);
        if ff.QuoteAction <> qaNone then
          WriteInteger(str, QIW_DF_QUOTE_ACTION, Integer(ff.QuoteAction));
        if ff.CharCase <> iccNone then
          WriteInteger(str, QIW_DF_CHAR_CASE, Integer(ff.CharCase));
        if ff.CharSet <> icsNone then
          WriteInteger(str, QIW_DF_CHAR_SET, Integer(ff.CharSet));
        // Replacements
        str := QIW_REPLACEMENTS + AnsiUpperCase(ff.FieldName);
        for j := 0 to ff.Replacements.Count - 1 do
        begin
          WriteString(Format('%s_%d', [str, j]), QIW_RP_TEXT_TO_FIND,
            ff.Replacements[j].TextToFind);
          WriteString(Format('%s_%d', [str, j]), QIW_RP_REPLACE_WITH,
            ff.Replacements[j].ReplaceWith);
          WriteBool(Format('%s_%d', [str, j]), QIW_RP_IGNORE_CASE,
            ff.Replacements[j].IgnoreCase);
        end;
      end;
      // Last Step
      WriteBool(QIW_LAST_STEP, QIW_COMMIT_AFTER_DONE, CommitAfterDone);
      WriteInteger(QIW_LAST_STEP, QIW_COMMIT_REC_COUNT, CommitRecCount);
      WriteInteger(QIW_LAST_STEP, QIW_IMPORT_REC_COUNT, ImportRecCount);
      WriteBool(QIW_LAST_STEP, QIW_CLOSE_AFTER_IMPORT, CloseAfterImport);
      WriteBool(QIW_LAST_STEP, QIW_ENABLE_ERROR_LOG, EnableErrorLog);
      WriteString(QIW_LAST_STEP, QIW_ERROR_LOG_FILE_NAME, ErrorLogFileName);
      WriteBool(QIW_LAST_STEP, QIW_REWRITE_ERROR_LOG_FILE, RewriteErrorLogFile);
      WriteBool(QIW_LAST_STEP, QIW_SHOW_ERROR_LOG, ShowErrorLog);
      WriteBool(QIW_LAST_STEP, QIW_AUTO_TRIM_VALUE, AutoTrimValue);
      WriteBool(QIW_LAST_STEP, QIW_IMPORT_EMPTY_ROWS, ImportEmptyRows);

      WriteInteger(QIW_LAST_STEP, QIW_IMPORT_MODE, Integer(ImportMode));
      WriteInteger(QIW_LAST_STEP, QIW_ADD_TYPE, Integer(AddType));
      EraseSection(QIW_KEY_COLUMNS);
      for i := 0 to lvSelectedColumns.Items.Count -1 do
        WriteString(QIW_KEY_COLUMNS, lvSelectedColumns.Items[i].Caption, EmptyStr);
    end;
  finally
    Template.Free;
  end;
end;

procedure TQImport3WizardF.btnLoadTemplateClick(Sender: TObject);
begin
  if odTemplate.Execute then begin
    if Assigned(Wizard.OnLoadTemplate) then
      Wizard.OnLoadTemplate(Wizard, odTemplate.FileName);
    LoadTemplateFromFile(odTemplate.FileName);
    Wizard.TemplateFileName := odTemplate.FileName;
  end;
end;

procedure TQImport3WizardF.btnSaveTemplateClick(Sender: TObject);
begin
  if sdTemplate.Execute then
    SaveTemplateToFile(sdTemplate.FileName);
end;

// XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS
// XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS
// XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS XLS
procedure TQImport3WizardF.XLSFillFieldList;
var
  i: Integer;
begin
  if not (aiXLS in Wizard.AllowedImports) then Exit;
  if not QImportDestinationAssigned(False, ImportDestination, DataSet, DBGrid,
           ListView, StringGrid) then Exit;
  XLSClearFieldList;
  for i := 0 to QImportDestinationColCount(False, ImportDestination, DataSet,
                  DBGrid, ListView, StringGrid) - 1 do
    with lvXLSFields.Items.Add do begin
      Caption := QImportDestinationColName(False, ImportDestination, DataSet,
        DBGrid, Self.ListView, StringGrid, GridCaptionRow, i);
      ImageIndex := 0;
      Data := TMapRow.Create(nil);
    end;
end;

procedure TQImport3WizardF.XLSClearFieldList;
var
  i, k: Integer;
  m: TMapRow;
begin
  for i := lvXLSFields.Items.Count - 1 downto 0 do begin
    if Assigned(lvXLSFields.Items[i].Data) then
    begin
      m := TMapRow(lvXLSFields.Items[i].Data);
      for k := m.Count-1 downto 0 do
        m.Delete(k);
      m.Free;
    end;
    lvXLSFields.Items.Delete(i);
  end;
end;

procedure TQImport3WizardF.XLSClearDataSheets;
var
  i: Integer;
begin
  for i := pcXLSFile.PageCount - 1 downto 0 do
    pcXLSFile.Pages[i].Free;
end;

procedure TQImport3WizardF.XLSFillGrid;
var
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  k, i, j, n: Integer;
  Cell: TbiffCell;
//  V: Variant;
  F: TForm;
  Start, Finish: TDateTime;
  W: Integer;
//  ExprLen: word;
//  Expr: PByteArray;
begin
  XLSClearDataSheets;

  if not FileExists(FFileName) then Exit;

  FXLSFile.FileName := FFileName;
  Start := Now;
  F := ShowLoading(Self, FFileName);
  try
    Application.ProcessMessages;
    FXLSFile.Clear;
    FXLSFile.Load;

    for k := 0 to FXLSFile.Workbook.WorkSheets.Count - 1 do
    begin
      TabSheet := TTabSheet.Create(pcXLSFile);
      TabSheet.PageControl := pcXLSFile;
      TabSheet.Caption := FXLSFile.Workbook.WorkSheets[k].Name;

      StringGrid := TqiStringGrid.Create(TabSheet);
      StringGrid.Parent := TabSheet;
      StringGrid.Align := alClient;
      StringGrid.ColCount := 257;
      n := 65536;
      if (Wizard.ExcelViewerRows > 0) and (Wizard.ExcelViewerRows <= 65536) then
        n := Wizard.ExcelViewerRows;
      StringGrid.RowCount := n + 1;
      StringGrid.FixedCols := 1;
      StringGrid.FixedRows := 1;
      StringGrid.DefaultColWidth := 64;
      StringGrid.DefaultRowHeight := 16;
      StringGrid.ColWidths[0] := 30;
      StringGrid.Options := StringGrid.Options - [goRangeSelect];
      StringGrid.OnClick := XLSGridClick;
      StringGrid.OnDrawCell := XLSDrawCell;
      StringGrid.OnMouseDown := XLSMouseDown;
      StringGrid.OnSelectCell := XLSSelectCell;
      StringGrid.OnExit := XLSGridExit;
      StringGrid.OnKeyDown := XLSGridKeyDown;
      StringGrid.Tag := 1000;

      GridFillFixedCells(StringGrid);

      for i := 0 to FXLSFile.Workbook.WorkSheets[k].Rows.Count - 1 do
        for j := 0 to FXLSFile.Workbook.WorkSheets[k].Rows[i].Count - 1 do begin
          Cell := FXLSFile.Workbook.WorkSheets[k].Rows[i][j];
          if (Cell.Col < StringGrid.ColCount - 1) and
             (Cell.Row < StringGrid.RowCount - 1) then begin
            case Cell.CellType of
              bctString  :
                StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] := Cell.AsString;
              bctBoolean :
                if Cell.AsBoolean
                  then StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] := 'True'
                  else StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] := 'False';
              bctNumeric :
                StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] :=
                  FloatToStr(Cell.AsFloat);
              bctDateTime:
                if Cell.AsDateTime = 0 then begin
                  StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] :=
                    FormatDateTime(FShortTimeFormat, Cell.AsDateTime)
                end
                else begin
                  if Cell.IsDateOnly then
                    StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] :=
                      FormatDateTime(FShortDateFormat, Cell.AsDateTime)
                  else if Cell.IsTimeOnly then
                    StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] :=
                      FormatDateTime(FShortTimeFormat, Cell.AsDateTime)
                  else StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] :=
                    FormatDateTime(FShortDateFormat + ' ' + FShortTimeFormat,
                      Cell.AsDateTime);
                end;
              bctUnknown :
(*  dee test
                if Cell.IsFormula then begin
                  ExprLen := GetWord(Cell.Data, 20);
                  if ExprLen > 0 then begin
                    GetMem(Expr, ExprLen);
                    try
                      Move(Cell.Data[22], Expr^, ExprLen);
                      V := CalculateFormula(Cell, Expr, ExprLen);
                    finally
                      FreeMem(Expr);
                    end;
                  end
                  else V := NULL;
                  StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] := VarToStr(V);
                end
                else
*)
                  StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] :=
                    VarToStr(Cell.AsVariant);
            end;
            W := StringGrid.Canvas.TextWidth(StringGrid.Cells[Cell.Col + 1, Cell.Row + 1]);
            if W + 10 > StringGrid.ColWidths[Cell.Col + 1] then
              if W + 10 < Wizard.ExcelMaxColWidth
                then StringGrid.ColWidths[Cell.Col + 1] := W + 10
                else StringGrid.ColWidths[Cell.Col + 1] := Wizard.ExcelMaxColWidth;
          end;
        end;
    end;
  finally
    Finish := Now;
    while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
      Finish := Now;

    if Assigned(F) then
      F.Free;
  end;

  FNeedLoadFile := False;
end;

procedure TQImport3WizardF.XLSTune;
begin
  ShowTip(tsExcelOptions, 6, 21, 39, tsExcelOptions.Width - 6, QImportLoadStr(QIW_XLS_Tip));

  pgImport.ActivePage := tsExcelOptions;
  laStep_04.Caption := Format(QImportLoadStr(QIW_Step), [1,3]);

  edXLSSkipCols.OnChange(nil);
  edXLSSkipRows.OnChange(nil);

  if lvXLSFields.Items.Count > 0 then
    if not Assigned(lvXLSFields.ItemFocused) then
    begin
      lvXLSFields.Items[0].Focused := True;
      lvXLSFields.Items[0].Selected := True;
      lvXLSFields.SetFocus;
    end;

  if FNeedLoadFile then XLSFillGrid;
  TuneButtons;
end;

procedure TQImport3WizardF.SetXLSSkipCols(const Value: Integer);
begin
  if FXLSSkipCols <> Value then begin
    FXLSSkipCols := Value;
    edXLSSkipCols.Text := IntToStr(Value);
    XLSRepaintCurrentGrid;
    TuneButtons;
  end;
end;

procedure TQImport3WizardF.SetXLSSkipRows(const Value: Integer);
begin
  if FXLSSkipRows <> Value then begin
    FXLSSkipRows := Value;
    edXLSSkipRows.Text := IntToStr(Value);
    XLSRepaintCurrentGrid;
    TuneButtons;
  end;
end;

function TQImport3WizardF.XLSReady: boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to lvXLSFields.Items.Count - 1 do
    if TMapRow(lvXLSFields.Items[i].Data).Count > 0 then begin
      Result := True;
      Break;
    end;
end;

procedure TQImport3WizardF.XLSDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  i: Integer;
begin
  FXLSDefinedRanges.Clear;
  if lvXLSFields.Focused then
  begin
    if Assigned(lvXLSFields.ItemFocused) and
       Assigned(lvXLSFields.ItemFocused.Data) then
      for i := 0 to TMapRow(lvXLSFields.ItemFocused.Data).Count - 1 do
        FXLSDefinedRanges.Add(TMapRow(lvXLSFields.ItemFocused.Data)[i]);
  end
  else begin
    if Assigned(lvXLSFields.ItemFocused) and
       Assigned(lvXLSFields.ItemFocused.Data) and
       Assigned(lvXLSRanges.ItemFocused) then
      FXLSDefinedRanges.Add(TMapRow(lvXLSFields.ItemFocused.Data)[lvXLSRanges.ItemFocused.Index]);
  end;

  if Sender is TqiStringGrid then
    GridDrawCell(Sender as TqiStringGrid, pcXLSFile.ActivePage.Caption,
      pcXLSFile.ActivePage.PageIndex + 1, ACol, ARow, Rect, State,
      FXLSDefinedRanges, XLSSkipCols, XLSSkipRows, FXLSIsEditingGrid,
      FXLSGridSelection);
end;

procedure TQImport3WizardF.XLSMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

  procedure AddColRowToSelection(IsCol, IsCtrl: boolean; Number: Integer);
  var
//    str, str1: AnsiString;
    str, str1: string;
    N: Integer;
  begin
    if IsCol
      then str := Format('[%s]%s-%s;', [pcXLSFile.ActivePage.Caption, Col2Letter(Number), COLFINISH])
      else str := Format('[%s]%s-%s;', [pcXLSFile.ActivePage.Caption, Row2Number(Number), ROWFINISH]);

      N := FXLSGridSelection.IndexOfRange(str);
      if N > -1 then FXLSGridSelection.Delete(N);

      if (not IsCtrl) or (N = -1) then begin
        str1 := FXLSGridSelection.AsString;
        str1 := str1 + str;
        FXLSGridSelection.AsString := str1;
      end;
  end;

var
  Grid: TqiStringGrid;
  ACol, ARow, SCol, SRow, N, i: Integer;
  IsCtrl, IsShift: boolean;

  procedure ChangeCurrentCell(Col, Row: Integer);
  var
    Event: TSelectCellEvent;
  begin
    Event := Grid.OnSelectCell;
    Grid.OnSelectCell := nil;
    Grid.Col := Col;
    Grid.Row := Row;
    Grid.OnSelectCell := Event;
  end;

begin
  if not (Sender is TqiStringGrid) then Exit;
  Grid := Sender as TqiStringGrid;

  IsShift := GetKeyState(VK_SHIFT) < 0;
  IsCtrl := GetKeyState(VK_CONTROL) < 0;

  if not (IsShift or IsCtrl) then
  begin
    Grid.Repaint;
    Exit;
  end;

  Grid.MouseToCell(X, Y, ACol, ARow);

  if not ((ACol = 0) xor (ARow = 0)) then
  begin
    Grid.Repaint;
    Exit;
  end;

  if not FXLSIsEditingGrid then
    XLSStartEditing;

  if IsCtrl then
  begin
    if ACol = 0
      then N := ARow
      else N := ACol;

    AddColRowToSelection(ARow = 0, True, N);

    if ACol = 0
      then ChangeCurrentCell(Grid.Col, ARow)
      else ChangeCurrentCell(ACol, Grid.Row);
  end
  else if IsShift then begin
    SCol := Grid.Col;
    SRow := Grid.Row;

    if ACol = 0 then begin
      if SRow <= ARow then
        for i := SRow to ARow do
          AddColRowToSelection(False, False, i)
      else
        for i := SRow downto ARow do
          AddColRowToSelection(False, False, i);
      ChangeCurrentCell(Grid.Col, ARow);
    end
    else begin
      if SCol <= ACol then
        for i := SCol to ACol do
          AddColRowToSelection(True, False, i)
      else
        for i := SCol downto ACol do
          AddColRowToSelection(True, False, i);
      ChangeCurrentCell(ACol, Grid.Row);
    end;
  end;

  XLSFillSelection;
  Grid.Repaint;
end;

procedure TQImport3WizardF.XLSSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
var
  Grid: TqiStringGrid;
  SCol, SRow, i: Integer;
//  Str: AnsiString;
  Str: string;
  IsShift, IsCtrl, Cut: boolean;
begin
  if not (Sender is TqiStringGrid) then Exit;
  Grid := Sender as TqiStringGrid;
//  try
    IsShift := GetKeyState(VK_SHIFT) < 0;
    IsCtrl := GetKeyState(VK_CONTROL) < 0;

    if not (IsShift or IsCtrl) then
    begin
      XLSFinishEditing;
      Exit;
    end;

    SCol := Grid.Col;
    SRow := Grid.Row;

    if IsShift and not ((SCol = ACol) or (SRow = ARow)) then
    begin
      XLSFinishEditing;
      Exit;
    end;

    if not FXLSIsEditingGrid then
      XLSStartEditing;

    Cut := False;
    if (SCol = ACol) and (SRow = ARow) then
    begin
    end
    else begin
      if IsShift then
      begin
        if FXLSGridSelection.Count > 0 then
        begin
          if SCol <> ACol then
          begin
            if FXLSGridSelection[FXLSGridSelection.Count - 1].RangeType = rtRow then
            begin
              if SCol > ACol then
              begin
                for i := SCol downto ACol + 1 do
                  if CellInRange(FXLSGridSelection[FXLSGridSelection.Count - 1],
                       pcXLSFile.ActivePage.Caption,
                       pcXLSFile.ActivePage.PageIndex + 1, i, ARow) then
                  begin
                    Cut := True;
                    FXLSGridSelection[FXLSGridSelection.Count - 1].Col2 :=
                      FXLSGridSelection[FXLSGridSelection.Count - 1].Col2 - 1;
                  end;
              end
              else begin
                for i := SCol to ACol - 1 do
                  if (FXLSGridSelection[FXLSGridSelection.Count - 1].RangeType = rtRow) and
                     CellInRange(FXLSGridSelection[FXLSGridSelection.Count - 1],
                       pcXLSFile.ActivePage.Caption,
                       pcXLSFile.ActivePage.PageIndex + 1, i, ARow) then
                  begin
                    Cut := True;
                    FXLSGridSelection[FXLSGridSelection.Count - 1].Col2 :=
                      FXLSGridSelection[FXLSGridSelection.Count - 1].Col2 + 1;
                  end;
              end;
            end
          end
          else if SRow <> ARow then
          begin
            if (FXLSGridSelection[FXLSGridSelection.Count - 1].RangeType = rtCol) then
            begin
              if SRow > ARow then
              begin
                for i := SRow downto ARow + 1 do
                  if CellInRange(FXLSGridSelection[FXLSGridSelection.Count - 1],
                       pcXLSFile.ActivePage.Caption,
                       pcXLSFile.ActivePage.PageIndex + 1, ACol, i) then
                  begin
                    Cut := True;
                    FXLSGridSelection[FXLSGridSelection.Count - 1].Row2 :=
                      FXLSGridSelection[FXLSGridSelection.Count - 1].Row2 - 1;
                  end;
              end
              else begin
                for i := SRow to ARow - 1 do
                  if CellInRange(FXLSGridSelection[FXLSGridSelection.Count - 1],
                       pcXLSFile.ActivePage.Caption,
                       pcXLSFile.ActivePage.PageIndex + 1, ACol, i) then
                  begin
                    Cut := True;
                    FXLSGridSelection[FXLSGridSelection.Count - 1].Row2 :=
                      FXLSGridSelection[FXLSGridSelection.Count - 1].Row2 + 1;
                  end;
              end;
            end;
          end;

          if (FXLSGridSelection[FXLSGridSelection.Count - 1].Col2 = SCol) and
             (FXLSGridSelection[FXLSGridSelection.Count - 1].Row2 = SRow) then
          begin
            if FXLSGridSelection[FXLSGridSelection.Count - 1].Col2 = ACol then
            begin
              if ARow > SRow
                then SRow := SRow + 1
                else SRow := SRow - 1;
            end
            else if FXLSGridSelection[FXLSGridSelection.Count - 1].Row2 = ARow then
            begin
              if ACol > SCol
                then SCol := SCol + 1
                else SCol := SCol - 1;
            end;
          end;

        end;
        if not Cut then
        begin
          str := FXLSGridSelection.AsString;
          if SCol = ACol then
          begin
            if SRow > ARow then
            begin
              for i := SRow downto ARow do
                if not CellInRow(FXLSGridSelection, EmptyStr, pcXLSFile.ActivePage.PageIndex + 1, ACol, i) then
                  str := str + Format('[%s]%s%d;', [pcXLSFile.ActivePage.Caption, Col2Letter(ACol), i]);
            end
            else begin
              for i := SRow to ARow do
                if not CellInRow(FXLSGridSelection, EmptyStr, pcXLSFile.ActivePage.PageIndex + 1, ACol, i) then
                  str := str + Format('[%s]%s%d;', [pcXLSFile.ActivePage.Caption, Col2Letter(ACol), i]);
            end
          end
          else if SRow = ARow then
          begin
            if SCol > ACol then
            begin
              for i := SCol downto ACol do
                if not CellInRow(FXLSGridSelection, EmptyStr, pcXLSFile.ActivePage.PageIndex + 1, i, ARow) then
                  str := str + Format('[%s]%s%d;', [pcXLSFile.ActivePage.Caption, Col2Letter(i), ARow]);
            end
            else begin
              for i := SCol to ACol do
                if not CellInRow(FXLSGridSelection, EmptyStr, pcXLSFile.ActivePage.PageIndex + 1, i, ARow) then
                  str := str + Format('[%s]%s%d;', [pcXLSFile.ActivePage.Caption, Col2Letter(i), ARow]);
            end
          end;
          FXLSGridSelection.AsString := str;
        end;
        XLSFillSelection;
      end
      else if IsCtrl then
      begin
        if not CellInRow(FXLSGridSelection, EmptyStr,
             pcXLSFile.ActivePage.PageIndex + 1, ACol, ARow) then
        begin
          str := FXLSGridSelection.AsString;
          str := str + Format('[%s]%s%d;', [pcXLSFile.ActivePage.Caption, Col2Letter(ACol), ARow]);
          FXLSGridSelection.AsString := str;
        end
        else begin
          RemoveCellFromRow(FXLSGridSelection, EmptyStr,
            pcXLSFile.ActivePage.PageIndex + 1, ACol, ARow);
        end;
        XLSFillSelection;
      end;
    end;
{  finally
    Grid.Repaint;
  end;}
end;

procedure TQImport3WizardF.XLSGridExit(Sender: TObject);
begin
  XLSFinishEditing;
end;

procedure TQImport3WizardF.XLSGridKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Shift = [] then
    case Key of
      VK_RETURN: XLSApplyEditing;
      VK_ESCAPE: XLSFinishEditing;
    end;
end;

procedure TQImport3WizardF.XLSGridClick(Sender: TObject);
begin
  if Sender is TqiStringGrid then
    (Sender as TqiStringGrid).Repaint;
end;

procedure TQImport3WizardF.XLSStartEditing;
begin
  FXLSIsEditingGrid := True;
  lvXLSRanges.Visible := False;
  lvXLSSelection.Visible := True;
  tbXLSRanges.Visible := False;
end;

procedure TQImport3WizardF.XLSFinishEditing;
begin
  FXLSIsEditingGrid := False;
  FXLSGridSelection.Clear;
  XLSRepaintCurrentGrid;

  lvXLSSelection.Visible:= False;

  lvXLSRanges.Visible := True;
  tbXLSRanges.Visible := True;;
end;

procedure TQImport3WizardF.XLSApplyEditing;
var
  i: Integer;
begin
  if FXLSGridSelection.Count = 0 then Exit;

  XLSDeleteSelectedRanges;

  for i := 0 to FXLSGridSelection.Count - 1 do begin
    FXLSGridSelection[i].MapRow := TMapRow(lvXLSFields.ItemFocused.Data);
    TMapRow(lvXLSFields.ItemFocused.Data).Add(FXLSGridSelection[i]);
    with lvXLSRanges.Items.Add do begin
      Caption := FXLSGridSelection[i].AsString;
      Data := FXLSGridSelection[i];
      ImageIndex := 3;
    end;
  end;
  if (lvXLSRanges.Items.Count > 0) and
     not Assigned(lvXLSRanges.ItemFocused) then begin
    lvXLSRanges.Items[0].Focused := True;
    lvXLSRanges.Items[0].Selected := True;
  end;

  XLSFinishEditing;
end;

procedure TQImport3WizardF.XLSDeleteSelectedRanges;
var
  List: TList;
  i: Integer;
begin
  if lvXLSRanges.SelCount = 0 then Exit;

  lvXLSRanges.OnChange := nil;
  try
    List := TList.Create;
    try
      for i := 0 to lvXLSRanges.Items.Count - 1 do
        if lvXLSRanges.Items[i].Selected then
          List.Add(Pointer(i));
      for i := List.Count - 1 downto 0 do begin
        TMapRow(lvXLSFields.ItemFocused.Data).Delete(Integer(List[i]));
        lvXLSRanges.Items[Integer(List[i])].Delete;
        List.Delete(i);
      end;

      if (lvXLSRanges.Items.Count > 0) and Assigned(lvXLSRanges.ItemFocused) and
         not lvXLSRanges.ItemFocused.Selected then
        lvXLSRanges.ItemFocused.Selected := True;

    finally
      List.Free;
    end;
    TuneButtons;
  finally
    lvXLSRanges.OnChange := lvXLSRangesChange;
  end;
end;

function TQImport3WizardF.XLSGetCurrentGrid: TqiStringGrid;
var
  i: Integer;
begin
  Result := nil;
  if not Assigned(pcXLSFile.ActivePage) then Exit;
  for i := 0 to pcXLSFile.ActivePage.ComponentCount - 1 do
    if (pcXLSFile.ActivePage.Components[i] is TqiStringGrid) and
       ((pcXLSFile.ActivePage.Components[i] as TqiStringGrid).Tag  = 1000) then
    begin
      Result := pcXLSFile.ActivePage.Components[i] as TqiStringGrid;
      Break;
    end;
end;

procedure TQImport3WizardF.XLSRepaintCurrentGrid;
var
  Grid: TqiStringGrid;
begin
  Grid := XLSGetCurrentGrid;
  if Assigned(Grid) then Grid.Repaint;
end;

procedure TQImport3WizardF.XLSFillSelection;
var
  i: Integer;
begin
//  lvXLSSelection.Items.BeginUpdate;
  try
    lvXLSSelection.Items.Clear;
    FXLSGridSelection.Optimize;
    for i := 0 to FXLSGridSelection.Count - 1 do
      with lvXLSSelection.Items.Add do begin
        Caption := FXLSGridSelection[i].AsString;
        ImageIndex := 3;
      end;
    if lvXLSSelection.Items.Count > 0 then begin
      lvXLSSelection.Items[0].Focused := True;
      lvXLSSelection.Items[0].Selected := True;
    end
  finally
  //  lvXLSSelection.Items.EndUpdate;
  end;
end;

procedure TQImport3WizardF.btnXLSAutoFillColsClick(Sender: TObject);
var
  i, j: Integer;
  MapRow: TMapRow;
  MR: TMapRange;
begin
  j := pcXLSFile.ActivePage.TabIndex;

  for i := 0 to lvXLSFields.Items.Count - 1 do begin
    MapRow := TMapRow(lvXLSFields.Items[i].Data);
    MapRow.Clear;
    if i <= FXLSFile.Workbook.WorkSheets[j].ColCount - 1 then begin
      MR := TMapRange.Create(MapRow);
      MR.Col1 := FXLSFile.Workbook.WorkSheets[j].Cols[i].ColNumber + 1;
      MR.Col2 := MR.Col1;
      MR.Row1 := 0;
      MR.Row2 := 0;
      MR.SheetIDType := sitName;
      MR.SheetName := FXLSFile.Workbook.WorkSheets[j].Name;
      MR.SheetNumber := 0;
      MR.Direction := rdDown;
      MapRow.Add(MR);
    end;
  end;
  if Assigned(lvXLSFields.ItemFocused) then
    lvXLSFieldsChange(lvXLSFields, lvXLSFields.ItemFocused, ctState);
  TuneButtons;
end;

procedure TQImport3WizardF.btnXLSAutoFillRowsClick(Sender: TObject);
var
  i, j: Integer;
  MapRow: TMapRow;
  MR: TMapRange;
begin
  j := pcXLSFile.ActivePage.TabIndex;

  for i := 0 to lvXLSFields.Items.Count - 1 do begin
    MapRow := TMapRow(lvXLSFields.Items[i].Data);
    MapRow.Clear;
    if i <= FXLSFile.Workbook.WorkSheets[j].RowCount - 1 then begin
      MR := TMapRange.Create(MapRow);
      MR.Row1 := FXLSFile.Workbook.WorkSheets[j].Rows[i].RowNumber + 1;
      MR.Row2 := MR.Row1;
      MR.Col1 := 0;
      MR.Col2 := 0;
      MR.SheetIDType := sitName;
      MR.SheetName := FXLSFile.Workbook.WorkSheets[j].Name;
      MR.SheetNumber := 0;
      MR.Direction := rdDown;
      MapRow.Add(MR);
    end;
  end;

  if Assigned(lvXLSFields.ItemFocused) then
    lvXLSFieldsChange(lvXLSFields, lvXLSFields.ItemFocused, ctState);
  TuneButtons;
end;

procedure TQImport3WizardF.lvXLSFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
var
  i: Integer;
  Row: TMapRow;
begin
  if csDestroying in ComponentState then Exit;

  if not Assigned(Item) or not Assigned(Item.Data) then Exit;
//  lvXLSRanges.Items.BeginUpdate;
  try
    lvXLSRanges.Items.Clear;
    Row := TMapRow(Item.Data);
    for i := 0 to Row.Count - 1 do
      with lvXLSRanges.Items.Add do begin
        Caption := Row[i].AsString;
        ImageIndex := 3;
        Data := Row[i];
      end;
    if lvXLSRanges.Items.Count > 0 then begin
      lvXLSRanges.Items[0].Focused := True;
      lvXLSRanges.Items[0].Selected := True;
    end
    else lvXLSRangesChange(lvXLSRanges, nil, ctState);

    XLSRepaintCurrentGrid;
  finally
  //  lvXLSRanges.Items.EndUpdate;
  end;
  TuneButtons;
end;

{procedure TQImportWizardF.lvXLSFieldsSelectItem(Sender: TObject;
  Item: TListItem; Selected: Boolean);
var
  i: Integer;
  Row: TMapRow;
begin
  if not Assigned(Item) or not Assigned(Item.Data) then Exit;
//  lvXLSRanges.Items.BeginUpdate;
  try
    lvXLSRanges.Items.Clear;
    Row := TMapRow(Item.Data);
    for i := 0 to Row.Count - 1 do
      with lvXLSRanges.Items.Add do begin
        Caption := Row[i].AsString;
        ImageIndex := 3;
        Data := Row[i];
      end;
    if lvXLSRanges.Items.Count > 0 then begin
      lvXLSRanges.Items[0].Focused := True;
      lvXLSRanges.Items[0].Selected := True;
    end
    else lvXLSRangesSelectItem(lvXLSRanges, nil, False);

    XLSRepaintCurrentGrid;
  finally
  //  lvXLSRanges.Items.EndUpdate;
  end;
  TuneButtons;
end;}

procedure TQImport3WizardF.tbtXLSAddRangeClick(Sender: TObject);
var
  Range: TMapRange;
  MapRow: TMapRow;
  i: Integer;
  Item: TListItem;
begin
  MapRow := TMapRow(lvXLSFields.ItemFocused.Data);
  Range := TMapRange.Create(MapRow);
  Range.Col1 := 1;
  Range.Row1 := 0;
  Range.Col2 := 1;
  Range.Row2 := 0;
  Range.Direction := rdDown;
  Range.SheetIDType := sitName;
  Range.SheetName := pcXLSFile.ActivePage.Caption;
  Range.SheetNumber := pcXLSFile.ActivePage.PageIndex + 1;
  Range.Update;

  if EditRange(Range, FXLSFile) then begin
    MapRow.Add(Range);
    lvXLSRanges.Items.BeginUpdate;
    try
      Item := lvXLSRanges.Items.Add;
      with Item do begin
        Caption := Range.AsString;
        ImageIndex := 3;
        Data := Range;
      end;
      for i := 0 to lvXLSRanges.Items.Count - 1 do begin
        lvXLSRanges.Items[i].Focused := lvXLSRanges.Items[i] = Item;
        lvXLSRanges.Items[i].Selected := lvXLSRanges.Items[i] = Item;
      end;
    finally
      lvXLSRanges.Items.EndUpdate;
    end;
  end
  else Range.Free;
  TuneButtons;
end;

procedure TQImport3WizardF.tbtXLSEditRangeClick(Sender: TObject);
begin
  if not Assigned(lvXLSFields.ItemFocused) and
     not Assigned(lvXLSRanges.ItemFocused) then Exit;

  if EditRange(TMapRange(lvXLSRanges.ItemFocused.Data), FXLSFile) then
    lvXLSRanges.ItemFocused.Caption := TMapRange(lvXLSRanges.ItemFocused.Data).AsString;
  XLSRepaintCurrentGrid;
  TuneButtons;
end;

procedure TQImport3WizardF.lvXLSRangesChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  XLSRepaintCurrentGrid;
  TuneButtons;
end;

procedure TQImport3WizardF.tbtXLSMoveRangeUpClick(Sender: TObject);
var
  Index, i: Integer;
begin
  Index := lvXLSRanges.ItemFocused.Index;
  TMapRow(lvXLSFields.ItemFocused.Data).Exchange(Index, Index - 1);
  lvXLSRanges.Items.BeginUpdate;
  try
    lvXLSRanges.Items[Index].Data := TMapRow(lvXLSFields.ItemFocused.Data)[Index];
    lvXLSRanges.Items[Index].Caption := TMapRange(lvXLSRanges.Items[Index].Data).AsString;
    lvXLSRanges.Items[Index - 1].Data := TMapRow(lvXLSFields.ItemFocused.Data)[Index - 1];
    lvXLSRanges.Items[Index - 1].Caption := TMapRange(lvXLSRanges.Items[Index - 1].Data).AsString;
    for i := 0 to lvXLSRanges.Items.Count - 1 do begin
      lvXLSRanges.Items[i].Focused := lvXLSRanges.Items[i] = lvXLSRanges.Items[Index - 1];
      lvXLSRanges.Items[i].Selected := lvXLSRanges.Items[i] = lvXLSRanges.Items[Index - 1];
    end;
  finally
    lvXLSRanges.Items.EndUpdate;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.tbtXLSMoveRangeDownClick(Sender: TObject);
var
  Index, i: Integer;
begin
  Index := lvXLSRanges.ItemFocused.Index;
  TMapRow(lvXLSFields.ItemFocused.Data).Exchange(Index, Index + 1);
  lvXLSRanges.Items.BeginUpdate;
  try
    lvXLSRanges.Items[Index].Data := TMapRow(lvXLSFields.ItemFocused.Data)[Index];
    lvXLSRanges.Items[Index].Caption := TMapRange(lvXLSRanges.Items[Index].Data).AsString;
    lvXLSRanges.Items[Index + 1].Data := TMapRow(lvXLSFields.ItemFocused.Data)[Index + 1];
    lvXLSRanges.Items[Index + 1].Caption := TMapRange(lvXLSRanges.Items[Index + 1].Data).AsString;
    for i := 0 to lvXLSRanges.Items.Count - 1 do begin
      lvXLSRanges.Items[i].Focused := lvXLSRanges.Items[i] = lvXLSRanges.Items[Index + 1];
      lvXLSRanges.Items[i].Selected := lvXLSRanges.Items[i] = lvXLSRanges.Items[Index + 1];
    end;
  finally
    lvXLSRanges.Items.EndUpdate;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.tbtXLSDelRangeClick(Sender: TObject);
begin
  XLSDeleteSelectedRanges;
  XLSRepaintCurrentGrid;
end;

procedure TQImport3WizardF.tbtXLSClearFieldRangesClick(Sender: TObject);
begin
  if not (Assigned(lvXLSFields.ItemFocused) and
          Assigned(lvXLSFields.ItemFocused.Data)) then Exit;
  TMapRow(lvXLSFields.ItemFocused.Data).Clear;
  lvXLSFieldsChange(lvXLSFields, lvXLSFields.ItemFocused, ctState);
end;

procedure TQImport3WizardF.edXLSSkipColsChange(Sender: TObject);
begin
  XLSSkipCols := StrToIntDef(edXLSSkipCols.Text, 0);
end;

procedure TQImport3WizardF.edXLSSkipRowsChange(Sender: TObject);
begin
  XLSSkipRows := StrToIntDef(edXLSSkipRows.Text, 0);
end;

procedure TQImport3WizardF.lvXLSRangesDblClick(Sender: TObject);
begin
  if tbtXLSEditRange.Enabled then
    tbtXLSEditRange.Click;
end;

procedure TQImport3WizardF.mScriptChange(Sender: TObject);
begin
  {$IFDEF USESCRIPT}
  if Assigned(lstFormatFields.Selected) and not FLoadingFormatItem then
  begin
    TQImportFieldFormat(lstFormatFields.Selected.Data).Script.Text :=
      mScript.Lines.Text;
    SetEnabledDataFormatControls;
  end;
  {$ENDIF}
end;

procedure TQImport3WizardF.tbtXLSClearAllRangesClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvXLSFields.Items.Count - 1 do
    TMapRow(lvXLSFields.Items[i].Data).Clear;
  lvXLSFieldsChange(lvXLSFields, lvXLSFields.ItemFocused, ctState);
  TuneButtons;
end;

{$IFDEF ADO}
// Access Access Access Access Access Access Access Access Access
// Access Access Access Access Access Access Access Access Access
// Access Access Access Access Access Access Access Access Access

procedure TQImport3WizardF.TuneAccessControls;
begin
  laAccessPassword.Enabled := FImportType = aiAccess;
  edAccessPassword.Enabled := FImportType = aiAccess;
end;

procedure TQImport3WizardF.SetAccessPassword(const Value: string);
begin
  if FAccessPassword <> Value then
  begin
    FAccessPassword := Value;
    if edAccessPassword.Enabled then
      edAccessPassword.Text := FAccessPassword;
  end;
end;

procedure TQImport3WizardF.SetAccessSourceType(
  const Value: TQImportAccessSourceType);
begin
  if FAccessSourceType <> Value then begin
    FAccessSourceType := Value;
    case FAccessSourceType of
      isTable: rbtAccessTable.Checked := true;
      isSQL: rbtAccessSQL.Checked := true;
    end;
    lbAccessTables.Enabled := FAccessSourceType = isTable;
    memAccessSQL.Enabled := FAccessSourceType = isSQL;
    tbtAccessSQLLoad.Enabled := FAccessSourceType = isSQL;
    tbtAccessSQLSave.Enabled := FAccessSourceType = isSQL;
    FNeedLoadFields := true;
    impAccess.SourceType := FAccessSourceType;
  end;
end;

procedure TQImport3WizardF.AccessClearList;
var
  i: integer;
begin
  for i := lstAccessDataSet.Items.Count - 1 downto 0 do
    lstAccessDataSet.Items.Delete(i);
end;

procedure TQImport3WizardF.AccessFillList;
var
  i: integer;
begin
  if not (aiAccess in Wizard.AllowedImports) then Exit;
  AccessClearList;
  for i := 0 to QImportDestinationColCount(false, ImportDestination, DataSet,
                  DBGrid, ListView, StringGrid) - 1 do
    lstAccessDataSet.Items.Add.Caption :=
      QImportDestinationColName(false, ImportDestination, DataSet, DBGrid,
        ListView, StringGrid, GridCaptionRow, i);
  if lstAccessDataSet.Items.Count > 0 then begin
    lstAccessDataSet.Items[0].Focused := true;
    lstAccessDataSet.Items[0].Selected := true;
  end;
end;

function TQImport3WizardF.AccessReady: boolean;
begin
  Result := lstAccessMap.Items.Count > 0;
end;

procedure TQImport3WizardF.AccessTune_01;
begin
  pgImport.ActivePage := tsAccessOptions_01;
  laStep_16.Caption := Format(QImportLoadStr(QIW_Step), [1,4]);

  if FNeedLoadFile then
    AccessGetTableNames;
  TuneButtons;
end;

procedure TQImport3WizardF.AccessTune_02;
begin
  ShowTip(tsAccessOptions_02, 6, 21, 39, tsAccessOptions_02.Width - 6, QImportLoadStr(QIW_Access_Tip));
  laStep_15.Caption := Format(QImportLoadStr(QIW_Step), [2,4]);
  if AccessSourceType = isTable then
  begin
    impAccess.TableName := lbAccessTables.Items[lbAccessTables.ItemIndex];
    impAccess.SQL.Clear;
    lstAccess.Column[0].Caption := impAccess.TableName;
    lstAccessMap.Column[2].Caption := impAccess.TableName;
  end
  else begin
    impAccess.TableName := EmptyStr;
    impAccess.SQL.Assign(memAccessSQL.Lines);
    lstAccess.Column[0].Caption := QImportLoadStr(QIW_Access_CustomQuery);
    lstAccessMap.Column[2].Caption := QImportLoadStr(QIW_Access_CustomQuery);
  end;
  AccessGetFieldNames;
  pgImport.ActivePage := tsAccessOptions_02;
  TuneButtons;
end;

procedure TQImport3WizardF.AccessGetTableNames;
var
  F: TForm;
begin
  impAccess.FileName := FFileName;
  impAccess.Password := AccessPassword;
  F := ShowLoading(Self, FFileName);
  try
    lbAccessTables.Items.Clear;
    Application.ProcessMessages;
    impAccess.GetTableNames(lbAccessTables.Items);
  finally
    F.Free;
  end;
  if Step = 1 then
  begin
    if AccessSourceType = isTable then
    begin
      if lbAccessTables.Enabled then
      begin
        ActiveControl := lbAccessTables;
        if lbAccessTables.Items.Count > 0 then
          lbAccessTables.ItemIndex := 0;
      end;
    end
    else begin
      if memAccessSQL.Enabled then
      begin
        ActiveControl := memAccessSQL;
        memAccessSQL.SetFocus;
      end;  
    end;
  end;
  FNeedLoadFile := false;
end;

procedure TQImport3WizardF.AccessGetFieldNames;
var
  FieldNames, FieldTypesString: TStringArray;
  FieldTypes: TFieldTypes;
  FieldSizes: TFieldSizes;
  i: integer;
  F: TForm;
begin
  if FNeedLoadFields then
  begin
    F := ShowLoading(Self, FFileName);
    try
      lstAccess.Items.BeginUpdate;
      try
        lstAccessMap.Items.BeginUpdate;
        try
          lstAccess.Items.Clear;
          lstAccessMap.Items.Clear;
          AccessFillList;
          Application.ProcessMessages;
          impAccess.GetFieldNames(FieldNames, FieldTypesString,
            FieldTypes, FieldSizes);
          for i := 0 to Length(FieldNames) - 1 do
            with lstAccess.Items.Add do Caption := FieldNames[i];
        finally
          lstAccessMap.Items.EndUpdate;
        end;
      finally
        lstAccess.Items.EndUpdate;
      end;
    finally
      F.Free;
    end;
    if lstAccess.Items.Count > 0 then begin
      lstAccess.Items[0].Focused := true;
      lstAccess.Items[0].Selected := true;
    end;
    FNeedLoadFields := false;
  end;
end;

procedure TQImport3WizardF.sbtAccessSQLLoadClick(Sender: TObject);
begin
  if odQuery.Execute then
    memAccessSQL.Lines.LoadFromFile(odQuery.FileName);
end;

procedure TQImport3WizardF.sbtAccessSQLSaveClick(Sender: TObject);
begin
  if sdQuery.Execute then
    memAccessSQL.Lines.SaveToFile(sdQuery.FileName);
end;

procedure TQImport3WizardF.lbAccessTablesClick(Sender: TObject);
begin
  FNeedLoadFields := true;
  TuneButtons;
end;

procedure TQImport3WizardF.memAccessSQLChange(Sender: TObject);
begin
  FNeedLoadFields := true;
  TuneButtons;
end;

procedure TQImport3WizardF.bAccessAddClick(Sender: TObject);
begin
  with lstAccessMap.Items.Add do begin
    Caption := lstAccessDataSet.Selected.Caption;
    SubItems.Add('=');
    SubItems.Add(lstAccess.Selected.Caption);
    ListView.Selected := lstAccessMap.Items[Index];
  end;
  lstAccessDataSet.Items.Delete(lstAccessDataSet.Selected.Index);
  if lstAccessDataSet.Items.Count > 0 then begin
    lstAccessDataSet.Items[0].Focused := true;
    lstAccessDataSet.Items[0].Selected := true;
  end;
  if lstAccessMap.Items.Count > 0 then begin
    lstAccessMap.Items[0].Focused := true;
    lstAccessMap.Items[0].Selected := true;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.bAccessAutoFillClick(Sender: TObject);
var
  i, N: integer;
begin
  lstAccessDataSet.Items.BeginUpdate;
  try
    lstAccess.Items.BeginUpdate;
    try
      lstAccessMap.Items.BeginUpdate;
      try
        lstAccessMap.Items.Clear;
        AccessClearList;
        AccessFillList;
        N := lstAccess.Items.Count;
        if N > lstAccessDataSet.Items.Count
          then N := lstAccessDataSet.Items.Count;
        for i := N - 1 downto 0 do begin
          with lstAccessMap.Items.Insert(0) do begin
            Caption := lstAccessDataSet.Items[i].Caption;
            SubItems.Add('=');
            SubItems.Add(lstAccess.Items[i].Caption);
          end;
          lstAccessDataSet.Items[i].Delete;
        end;
        if lstAccessMap.Items.Count > 0
          then lstAccessMap.Items[0].Selected := true;
      finally
        lstAccessMap.Items.EndUpdate;
      end;
    finally
      lstAccess.Items.EndUpdate;
    end;
  finally
    lstAccessDataSet.Items.EndUpdate;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.bAccessRemoveClick(Sender: TObject);
begin
  lstAccessDataSet.Items.Add.Caption := lstAccessMap.Selected.Caption;
  lstAccessMap.Items.Delete(lstAccessMap.Selected.Index);
  if lstAccessMap.Items.Count > 0 then begin
    lstAccessMap.Items[0].Focused := true;
    lstAccessMap.Items[0].Selected := true;
  end;
  if lstAccessDataSet.Items.Count > 0 then begin
    lstAccessDataSet.Items[0].Focused := true;
    lstAccessDataSet.Items[0].Selected := true;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.bAccessClearClick(Sender: TObject);
begin
  lstAccessDataSet.Items.BeginUpdate;
  try
    lstAccessMap.Items.BeginUpdate;
    try
      lstAccessMap.Items.Clear;
      AccessClearList;
      AccessFillList;
    finally
      lstAccessMap.Items.EndUpdate;
    end;
  finally
    lstAccessDataSet.Items.EndUpdate;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.lstAccessMapChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  TuneButtons;
end;

procedure TQImport3WizardF.lstAccessDataSetChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  TuneButtons;
end;

procedure TQImport3WizardF.lstAccessChange(Sender: TObject; Item: TListItem;
  Change: TItemChange);
begin
  TuneButtons;
end;

procedure TQImport3WizardF.edAccessPasswordChange(Sender: TObject);
begin
  AccessPassword := edAccessPassword.Text;
end;

procedure TQImport3WizardF.rbtAccessTableClick(Sender: TObject);
begin
  AccessSourceType := isTable;
  TuneButtons;
end;

procedure TQImport3WizardF.rbtAccessSQLClick(Sender: TObject);
begin
  AccessSourceType := isSQL;
  TuneButtons;
end;

procedure TQImport3WizardF.pbAccessAddPaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := true;
    if bAccessAdd.Enabled
      then i := 4
      else i := 12;
    ilWizard.GetBitmap(i, Bmp);
    pbAccessAdd.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TQImport3WizardF.bAccessAddMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAccessAdd.Left := pbAccessAdd.Left + 1;
  pbAccessAdd.Top := pbAccessAdd.Top + 1;
end;

procedure TQImport3WizardF.bAccessAddMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAccessAdd.Left := pbAccessAdd.Left - 1;
  pbAccessAdd.Top := pbAccessAdd.Top - 1;
end;

procedure TQImport3WizardF.pbAccessAutoFillPaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := true;
    if bAccessAutoFill.Enabled
      then i := 9
      else i := 17;
    ilWizard.GetBitmap(i, Bmp);
    pbAccessAutoFill.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TQImport3WizardF.bAccessAutoFillMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAccessAutoFill.Left := pbAccessAutoFill.Left + 1;
  pbAccessAutoFill.Top := pbAccessAutoFill.Top + 1;
end;

procedure TQImport3WizardF.bAccessAutoFillMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAccessAutoFill.Left := pbAccessAutoFill.Left - 1;
  pbAccessAutoFill.Top := pbAccessAutoFill.Top - 1;
end;

procedure TQImport3WizardF.pbAccessRemovePaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := true;
    if bAccessRemove.Enabled
      then i := 6
      else i := 14;
    ilWizard.GetBitmap(i, Bmp);
    pbAccessRemove.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TQImport3WizardF.bAccessRemoveMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAccessRemove.Left := pbAccessRemove.Left + 1;
  pbAccessRemove.Top := pbAccessRemove.Top + 1;
end;

procedure TQImport3WizardF.bAccessRemoveMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAccessRemove.Left := pbAccessRemove.Left - 1;
  pbAccessRemove.Top := pbAccessRemove.Top - 1;
end;

procedure TQImport3WizardF.pbAccessClearPaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := true;
    if bAccessClear.Enabled
      then i := 7
      else i := 15;
    ilWizard.GetBitmap(i, Bmp);
    pbAccessClear.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TQImport3WizardF.bAccessClearMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAccessClear.Left := pbAccessClear.Left + 1;
  pbAccessClear.Top := pbAccessClear.Top + 1;
end;

procedure TQImport3WizardF.bAccessClearMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAccessClear.Left := pbAccessClear.Left - 1;
  pbAccessClear.Top := pbAccessClear.Top - 1;
end;
{$ENDIF}

// Formats Formats Formats Formats Formats Formats Formats Formats Formats
// Formats Formats Formats Formats Formats Formats Formats Formats Formats
// Formats Formats Formats Formats Formats Formats Formats Formats Formats

procedure TQImport3WizardF.SetDecimalSeparator(const Value: Char);
begin
  if FDecimalSeparator <> Value then begin
    FDecimalSeparator := Value;
    edtDecimalSeparator.Text := Char2Str(FDecimalSeparator);
  end;
end;

procedure TQImport3WizardF.SetThousandSeparator(const Value: Char);
begin
  if FThousandSeparator <> Value then begin
    FThousandSeparator := Value;
    edtThousandSeparator.Text := Char2Str(FThousandSeparator);
  end;
end;

procedure TQImport3WizardF.SetShortDateFormat(const Value: string);
begin
  if FShortDateFormat <> Value then begin
    FShortDateFormat := Value;
    edtShortDateFormat.Text := FShortDateFormat;
  end;
end;

procedure TQImport3WizardF.SetLongDateFormat(const Value: string);
begin
  if FLongDateFormat <> Value then begin
    FLongDateFormat := Value;
    edtLongDateFormat.Text := FLongDateFormat;
  end;
end;

procedure TQImport3WizardF.SetDateSeparator(const Value: Char);
begin
  if FDateSeparator <> Value then begin
    FDateSeparator := Value;
    edtDateSeparator.Text := Char2Str(FDateSeparator);
  end;
end;

procedure TQImport3WizardF.SetShortTimeFormat(const Value: string);
begin
  if FShortTimeFormat <> Value then begin
    FShortTimeFormat := Value;
    edtShortTimeFormat.Text := FShortTimeFormat;
  end;
end;

procedure TQImport3WizardF.SetLongTimeFormat(const Value: string);
begin
  if FLongTimeFormat <> Value then begin
    FLongTimeFormat := Value;
    edtLongTimeFormat.Text := FlongTimeFormat;
  end;
end;

procedure TQImport3WizardF.SetTimeSeparator(const Value: Char);
begin
  if FTimeSeparator <> Value then begin
    FTimeSeparator := Value;
    edtTimeSeparator.Text := Char2Str(FTimeSeparator);
  end;
end;

procedure TQImport3WizardF.edtDecimalSeparatorExit(Sender: TObject);
begin
  DecimalSeparator := Str2Char(edtDecimalSeparator.Text, DecimalSeparator);
end;

procedure TQImport3WizardF.edtThousandSeparatorExit(Sender: TObject);
begin
  ThousandSeparator := Str2Char(edtThousandSeparator.Text, ThousandSeparator);
end;

procedure TQImport3WizardF.edtShortDateFormatChange(Sender: TObject);
begin
  ShortDateFormat := edtShortDateFormat.Text;
end;

procedure TQImport3WizardF.edtLongDateFormatChange(Sender: TObject);
begin
  LongDateFormat := edtLongDateFormat.Text;
end;

procedure TQImport3WizardF.edtDateSeparatorExit(Sender: TObject);
begin
  DateSeparator := Str2Char(edtDateSeparator.Text, DateSeparator);
end;

procedure TQImport3WizardF.edtShortTimeFormatChange(Sender: TObject);
begin
  ShortTimeFormat := edtShortTimeFormat.Text;
end;

procedure TQImport3WizardF.edtLongTimeFormatChange(Sender: TObject);
begin
  LongTimeFormat := edtLongTimeFormat.Text;
end;

procedure TQImport3WizardF.edtTimeSeparatorExit(Sender: TObject);
begin
  TimeSeparator := Str2Char(edtTimeSeparator.Text, TimeSeparator);
end;

procedure TQImport3WizardF.FormatsFillList;
var
  i, j: Integer;
begin
  FormatsClearList;
  for i := 0 to QImportDestinationColCount(False,
                  ImportDestination, DataSet, DBGrid, ListView,
                  StringGrid) - 1 do begin
    with lstFormatFields.Items.Add do begin
      Caption := QImportDestinationColName(False, ImportDestination, DataSet,
                 DBGrid, Self.ListView, StringGrid, GridCaptionRow, i);
      Data := TQImportFieldFormat.Create(nil);
      ImageIndex := 0;
      with TQImportFieldFormat(Data) do begin
        FieldName := QImportDestinationColName(False, ImportDestination,
                     DataSet, DBGrid, Self.ListView, StringGrid, GridCaptionRow, i);
        j := FieldFormats.IndexByName(FieldName);
        if j = -1 then Continue;
        GeneratorValue := FieldFormats[j].GeneratorValue;
        GeneratorStep := FieldFormats[j].GeneratorStep;
        ConstantValue := FieldFormats[j].ConstantValue;
        NullValue := FieldFormats[j].NullValue;
        DefaultValue := FieldFormats[j].DefaultValue;
        QuoteAction := FieldFormats[j].QuoteAction;
        LeftQuote := FieldFormats[j].LeftQuote;
        RightQuote := FieldFormats[j].RightQuote;
        CharCase := FieldFormats[j].CharCase;
        CharSet := FieldFormats[j].CharSet;
      end;
    end;
  end;
  if lstFormatFields.Items.Count > 0 then begin
    lstFormatFields.Items[0].Focused := True;
    lstFormatFields.Items[0].Selected := True;
    ShowFormatItem(lstFormatFields.Items[0])
  end;
end;

procedure TQImport3WizardF.FormatsClearList;
var
  i: Integer;
begin
  for i := lstFormatFields.Items.Count - 1 downto 0 do begin
    if Assigned(lstFormatFields.Items[i].Data) then
      TQImportFieldFormat(lstFormatFields.Items[i].Data).Free;
    lstFormatFields.Items.Delete(i);
  end;
end;

procedure TQImport3WizardF.ShowFormatItem(Item: TListItem);
var
  FieldFormat: TQImportFieldFormat;
  i: Integer;
begin
  if not Assigned(Item) then Exit;
  FLoadingFormatItem := True;
  try
    FieldFormat := TQImportFieldFormat(Item.Data);
    edtGeneratorValue.Text := IntToStr(FieldFormat.GeneratorValue);
    edtGeneratorStep.Text := IntToStr(FieldFormat.GeneratorStep);
    edtConstantValue.Text := FieldFormat.ConstantValue;
    edtNullValue.Text := FieldFormat.NullValue;
    edtDefaultValue.Text := FieldFormat.DefaultValue;
    edtLeftQuote.Text := FieldFormat.LeftQuote;
    edtRightQuote.Text := FieldFormat.RightQuote;
    cmbQuoteAction.ItemIndex := Integer(FieldFormat.QuoteAction);
    cmbCharCase.ItemIndex := Integer(FieldFormat.CharCase);
    cmbCharSet.ItemIndex := Integer(FieldFormat.CharSet);

    lvReplacements.Items.BeginUpdate;
    try
      lvReplacements.Items.Clear;
      for i := 0 to FieldFormat.Replacements.Count - 1 do
        with lvReplacements.Items.Add do
        begin
          Caption := FieldFormat.Replacements[i].TextToFind;
          SubItems.Add(FieldFormat.Replacements[i].ReplaceWith);
          if FieldFormat.Replacements[i].IgnoreCase
            then SubItems.Add(QImportLoadStr(QIWDF_IgnoreCase_Yes))
            else SubItems.Add(QImportLoadStr(QIWDF_IgnoreCase_No));
          ImageIndex := 3;
          Data := FieldFormat.Replacements[i];
        end;
      if lvReplacements.Items.Count > 0 then
      begin
        lvReplacements.Items[0].Focused := True;
        lvReplacements.Items[0].Selected := True;
      end;
    finally
      lvReplacements.Items.EndUpdate;
    end;

    SetEnabledDataFormatControls;
  finally
    FLoadingFormatItem := False;
  end;
end;

// Last Step Last Step Last Step Last Step Last Step Last Step Last Step Last Step
// Last Step Last Step Last Step Last Step Last Step Last Step Last Step Last Step
// Last Step Last Step Last Step Last Step Last Step Last Step Last Step Last Step

procedure TQImport3WizardF.SetCommitAfterDone(const Value: boolean);
begin
  if FCommitAfterDone <> Value then begin
    FCommitAfterDone := Value;
    chCommitAfterDone.Checked := FCommitAfterDone;
  end;
end;

procedure TQImport3WizardF.SetCommitRecCount(const Value: Integer);
begin
  if FCommitRecCount <> Value then begin
    FCommitRecCount := Value;
    edtCommitRecCount.Text := IntToStr(FCommitRecCount);
  end;
end;

procedure TQImport3WizardF.chCommitAfterDoneClick(Sender: TObject);
begin
  CommitAfterDone := chCommitAfterDone.Checked;
end;

procedure TQImport3WizardF.edtCommitRecCountChange(Sender: TObject);
begin
  try
    CommitRecCount := StrToInt(edtCommitRecCount.Text);
  except
  end;
end;

procedure TQImport3WizardF.SetImportRecCount(const Value: Integer);
begin
  if FImportRecCount <> Value then
    FImportRecCount := Value;
  edtImportRecCount.Text := IntToStr(FImportRecCount);
  chImportAllRecords.Checked := FImportRecCount = 0;
end;

procedure TQImport3WizardF.chImportAllRecordsClick(Sender: TObject);
begin
  edtImportRecCount.Enabled := not chImportAllRecords.Checked;
  laImportRecCount_01.Enabled := not chImportAllRecords.Checked;
  laImportRecCount_02.Enabled := not chImportAllRecords.Checked;
end;

procedure TQImport3WizardF.chImportEmptyRowsClick(Sender: TObject);
begin
  ImportEmptyRows := chImportEmptyRows.Checked;
end;

procedure TQImport3WizardF.edtImportRecCountChange(Sender: TObject);
begin
  try
    ImportRecCount := StrToInt(edtImportRecCount.Text);
  except
    edtImportRecCount.Text := IntToStr(ImportRecCount);
  end;
end;

procedure TQImport3WizardF.SetCloseAfterImport(const Value: boolean);
begin
  if FCloseAfterImport <> Value then
  begin
    FCloseAfterImport := Value;
    chCloseAfterImport.Checked := FCloseAfterImport;
  end;
end;

procedure TQImport3WizardF.chCloseAfterImportClick(Sender: TObject);
begin
  CloseAfterImport := chCloseAfterImport.Checked;
end;

procedure TQImport3WizardF.SetEnableErrorLog(const Value: boolean);
begin
  if FEnableErrorLog <> Value then begin
    FEnableErrorLog := Value;
    chEnableErrorLog.Checked := FEnableErrorLog;
  end;
end;

procedure TQImport3WizardF.SetErrorLogFileName(const Value: qiString);
begin
  if FErrorLogFileName <> Value then begin
    FErrorLogFileName := Value;
    edErrorLogFileName.Text := FErrorLogFileName;
  end;
end;

procedure TQImport3WizardF.SetRewriteErrorLogFile(const Value: boolean);
begin
  if FRewriteErrorLogFile <> Value then begin
    FRewriteErrorLogFile := Value;
    chRewriteErrorLogFile.Checked := FRewriteErrorLogFile;
  end;
end;

procedure TQImport3WizardF.SetShowErrorLog(const Value: boolean);
begin
  if FShowErrorLog <> Value then begin
    if Value and not EnableErrorLog then Exit;
    FShowErrorLog := Value;
    chShowErrorLog.Checked := FShowErrorLog;
  end;
end;

procedure TQImport3WizardF.SetImportEmptyRows(const Value: Boolean);
begin
  if FImportEmptyRows <> Value then
  begin
    FImportEmptyRows := Value;
    chImportEmptyRows.Checked := FImportEmptyRows;
  end;
end;

procedure TQImport3WizardF.SetImportMode(const Value: TQImportMode);
begin
  if FImportMode <> Value then begin
    FImportMode := Value;
    rgImportMode.ItemIndex := Integer(FImportMode);
  end;
end;

procedure TQImport3WizardF.SetAddType(const Value: TQImportAddType);
begin
  if FAddType <> Value then begin
    FAddType := Value;
    rgAddType.ItemIndex := Integer(FAddType);
  end;
end;

procedure TQImport3WizardF.chEnableErrorLogClick(Sender: TObject);
begin
  EnableErrorLog := chEnableErrorLog.Checked;
  chShowErrorLog.Enabled := FEnableErrorLog;
  laErrorLogFileName.Enabled := FEnableErrorLog;
  edErrorLogFileName.Enabled := FEnableErrorLog;
  chRewriteErrorLogFile.Enabled := FEnableErrorLog;
  bErrorLogFileName.Enabled := FEnableErrorLog;
end;

procedure TQImport3WizardF.edErrorLogFileNameChange(Sender: TObject);
begin
  ErrorLogFileName := edErrorLogFileName.Text;
end;

procedure TQImport3WizardF.bErrorLogFileNameClick(Sender: TObject);
var
  Path: string;
begin
  if edErrorLogFileName.Text <> EmptyStr then begin
    Path := ExtractFilePath(edErrorLogFileName.Text);
    if Path = EmptyStr then GetDir(0, Path);
    sdErrorLog.InitialDir := Path;
  end;
  if sdErrorLog.Execute then edErrorLogFileName.Text := sdErrorLog.FileName;
end;

procedure TQImport3WizardF.chRewriteErrorLogFileClick(Sender: TObject);
begin
  RewriteErrorLogFile := chRewriteErrorLogFile.Checked;
end;

procedure TQImport3WizardF.chShowErrorLogClick(Sender: TObject);
begin
  ShowErrorLog := chShowErrorLog.Checked;
end;

procedure TQImport3WizardF.rgImportModeClick(Sender: TObject);
begin
  ImportMode := TQImportMode(rgImportMode.ItemIndex);
end;

procedure TQImport3WizardF.rgAddTypeClick(Sender: TObject);
begin
  AddType := TQImportAddType(rgAddType.ItemIndex);
end;

procedure TQImport3WizardF.bAllToRightClick(Sender: TObject);
begin
  MoveToSelected(lvAvailableColumns, lvSelectedColumns, True, -1);
end;

procedure TQImport3WizardF.bOneToRirghtClick(Sender: TObject);
begin
  MoveToSelected(lvAvailableColumns, lvSelectedColumns, False, -1);
end;

procedure TQImport3WizardF.bOneToLeftClick(Sender: TObject);
begin
  MoveToSelected(lvSelectedColumns, lvAvailableColumns, False, -1);
end;

procedure TQImport3WizardF.bAllToLeftClick(Sender: TObject);
begin
  MoveToSelected(lvSelectedColumns, lvAvailableColumns, True, -1);
end;

procedure TQImport3WizardF.lvAvailableColumnsDblClick(Sender: TObject);
begin
  MoveToSelected(lvAvailableColumns, lvSelectedColumns, False, -1);
end;

procedure TQImport3WizardF.lvSelectedColumnsDblClick(Sender: TObject);
begin
  MoveToSelected(lvSelectedColumns, lvAvailableColumns, False, -1);
end;

procedure TQImport3WizardF.KeyColumnsDragOver(Sender,
  Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := Source is TListView;
end;

procedure TQImport3WizardF.KeyColumnsDragDrop(Sender,
  Source: TObject; X, Y: Integer);
var
  ListItem: TListItem;
  Index, i: Integer;
  str: string;
begin
  if Source <> Sender then
  begin
    ListItem := (Sender as TListView).GetItemAt(X, Y);
    if Assigned(ListItem)
      then Index := ListItem.Index + 1
      else Index := -1;
    if Source = lvAvailableColumns
      then MoveToSelected(lvAvailableColumns, lvSelectedColumns, False, Index);
    if Source = lvSelectedColumns
      then MoveToSelected(lvSelectedColumns, lvAvailableColumns, False, Index);
  end
  else begin
    with (Sender as TListView) do
    begin
      ListItem := GetItemAt(X, Y);
      if Assigned(ListItem) and Assigned(Selected) then
      begin
        Items.BeginUpdate;
        try
          str := Selected.Caption;
          Items.Delete(Selected.Index);
          ListItem := Items.Insert(ListItem.Index + 1);
          ListItem.Caption := str;
          for i := 0 to Items.Count - 1 do
          begin
            Items[i].Focused := Items[i] = ListItem;
            Items[i].Selected := Items[i] = ListItem;
          end;
        finally
          Items.EndUpdate;
        end;
      end;
    end;
  end;
end;

// TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT
// TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT
// TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT TXT

procedure TQImport3WizardF.TXTTune;
var
  i, j, k, p, s: Integer;
  strings: TStrings;
begin
  if FNeedLoadFile then
  begin
    vwTXT.LineCount := Wizard.TextViewerRows;
    vwTXT.LoadFromFile(FileName);
    FNeedLoadFile := False;

    cmbTXTEncoding.ItemIndex := Integer(ctWinDefined);

    if (FTmpFileName <> '') and FileExists(FileName) then
    begin
      with TIniFile.Create(FTmpFileName) do
      try
        strings := TStringList.Create;
        try
          ReadSection(QIW_TXT_MAP, strings);
          for i := 0 to strings.Count - 1 do
          begin
            j := -1;
            for k := 0 to lvTXTFields.Items.Count - 1 do
            begin
              if AnsiCompareText(lvTXTFields.Items[k].Caption, strings[i]) = 0 then
              begin
                j := k;
                Break;
              end;
            end;
            if j > -1 then
            begin
              TXTExtractPosSize(ReadString(QIW_TXT_MAP, strings[i],
                EmptyStr), p, s);
              vwTXT.AddArrow(p);
              vwTXT.AddArrow(p + s);
              lvTXTFields.Items[j].SubItems[0] := IntToStr(p);
              lvTXTFields.Items[j].SubItems[1] := IntToStr(s);
            end;
          end;
          TXTSkipLines := ReadInteger(QIW_TXT_OPTIONS, QIW_TXT_SKIP_LINES, 0);
          TXTEncoding := TQICharsetType(ReadInteger(QIW_TXT_OPTIONS, QIW_TXT_ENCODING_OPTION, 0));
        finally
          strings.Free;
        end;
      finally
        Free;
      end;
    end;
  end;

  ShowTip(tsTXTOptions, 6, 21, 39, tsTXTOptions.Width - 6, QImportLoadStr(QIW_TXT_Tip));

  pgImport.ActivePage := tsTXTOptions;
  laTXTStep_02.Caption := Format(QImportLoadStr(QIW_Step), [1, 3]);


  if (lvTXTFields.Items.Count > 0) and
     not Assigned(lvTXTFields.Selected) then
  begin
    lvTXTFields.Items[0].Focused := True;
    lvTXTFields.Items[0].Selected := True;
  end;

  lvTXTFieldsChange(lvTXTFields, lvTXTFields.Selected, ctState);
  TuneButtons;
end;

function TQImport3WizardF.TXTReady: boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to lvTXTFields.Items.Count - 1 do
  begin
    Result := StrToIntDef(lvTXTFields.Items[i].SubItems[1], -1) > -1;
    if Result then Break;
  end;
end;

procedure TQImport3WizardF.TXTFillCombo;
var
  i: Integer;
begin
  if not (aiTXT in Wizard.AllowedImports) then Exit;
  if not QImportDestinationAssigned(False, ImportDestination, DataSet, DBGrid,
           ListView, StringGrid) then Exit;
  TXTClearCombo;
  for i := 0 to QImportDestinationColCount(False, ImportDestination, DataSet,
                  DBGrid, ListView, StringGrid) - 1 do
    with lvTXTFields.Items.Add do
    begin
      Caption := QImportDestinationColName(False, ImportDestination, DataSet,
        DBGrid, Self.ListView, StringGrid, GridCaptionRow, i);
      SubItems.Add(EmptyStr);
      SubItems.Add(EmptyStr);
    end;
end;

procedure TQImport3WizardF.TXTClearCombo;
begin
  lvTXTFields.Items.Clear;
end;

procedure TQImport3WizardF.TXTExtractPosSize(const S: string; var Position,
  Size: Integer);
var
  i: Integer;
  b: boolean;
begin
  i := Pos(';', S);
  b := S[Length(S)] = ';';
  if i = 0 then Exit;
  try
    Position := StrToInt(Copy(S, 1, i - 1));
    Size := StrToInt(Copy(S, i + 1, Length(S) - i - Integer(b)))
  except
  end;
end;

procedure TQImport3WizardF.SetTXTEncoding(const Value: TQICharsetType);
begin
  if FTXTEncoding <> Value then
  begin
    FTXTEncoding := Value;
    cmbTXTEncoding.ItemIndex := Integer(FTXTEncoding);
  end;
end;

procedure TQImport3WizardF.SetCSVEncoding(const Value: TQICharsetType);
begin
  if FCSVEncoding <> Value then
  begin
    FCSVEncoding := Value;
    cmbCSVEncoding.ItemIndex := Integer(FCSVEncoding);
  end;
end;

procedure TQImport3WizardF.cmbTXTEncodingChange(Sender: TObject);
begin
  FTXTEncoding := TQICharsetType(cmbTXTEncoding.ItemIndex);
  vwTXT.CodePage := SystemCTNames[FTXTEncoding];
  vwTXT.LoadFromFile(FileName);
end;

procedure TQImport3WizardF.cmbCSVEncodingChange(Sender: TObject);
begin
  FCSVEncoding := TQICharsetType(cmbCSVEncoding.ItemIndex);

  CSVFillGrid;
end;

procedure TQImport3WizardF.SetTXTSkipLines(const Value: Integer);
begin
  if FTXTSkipLines <> Value then
  begin
    FTXTSkipLines := Value;
    edtTXTSkipLines.Text := IntToStr(FTXTSkipLines);
  end;
end;

procedure TQImport3WizardF.edtTXTSkipLinesChange(Sender: TObject);
begin
  TXTSkipLines := StrToIntDef(edtTXTSkipLines.Text, 0);
end;

procedure TQImport3WizardF.lvTXTFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
var
  P, S: Integer;
begin
  if not Assigned(Item) then Exit;
  if Item.SubItems.Count < 2 then Exit;
  P := StrToIntDef(Item.SubItems[0], -1);
  S := StrToIntDef(Item.SubItems[1], -1);
  vwTXT.SetSelection(P, S);
  vwTXT.Repaint;
  TuneButtons;
end;

procedure TQImport3WizardF.tbtTXTClearClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvTXTFields.Items.Count - 1 do
  begin
    lvTXTFields.Items[i].SubItems[0] := EmptyStr;
    lvTXTFields.Items[i].SubItems[1] := EmptyStr;
  end;
  lvTXTFieldsChange(lvTXTFields, lvTXTFields.ItemFocused, ctState);
end;

procedure TQImport3WizardF.TXTViewerChangeSelection(Sender: TObject);
var
  P, S: Integer;
begin
  if not Assigned(lvTXTFields.ItemFocused) then Exit;
  vwTXT.GetSelection(P, S);
  if (P > -1) and (S > -1) then begin
    lvTXTFields.ItemFocused.SubItems[0] := IntToStr(P);
    lvTXTFields.ItemFocused.SubItems[1] := IntToStr(S);
  end
  else begin
    lvTXTFields.ItemFocused.SubItems[0] := EmptyStr;
    lvTXTFields.ItemFocused.SubItems[1] := EmptyStr;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.TXTViewerDeleteArrow(Sender: TObject;
  Position: Integer);
var
  i, p, s: Integer;
begin
  for i := 0 to lvTXTFields.Items.Count - 1 do
  begin
    p := StrToIntDef(lvTXTFields.Items[i].SubItems[0], -1);
    s := StrToIntDef(lvTXTFields.Items[i].SubItems[1], -1);
    if (p > -1) and (s > -1) and
       ((p = Position) or (p + s = Position)) then
    begin
      lvTXTFields.Items[i].SubItems[0] := EmptyStr;
      lvTXTFields.Items[i].SubItems[1] := EmptyStr;
    end;
  end;
end;

procedure TQImport3WizardF.TXTViewerMoveArrow(Sender: TObject;
  OldPos, NewPos: Integer);
var
  i, p, s: Integer;
begin
  for i := 0 to lvTXTFields.Items.Count - 1 do
  begin
    p := StrToIntDef(lvTXTFields.Items[i].SubItems[0], -1);
    s := StrToIntDef(lvTXTFields.Items[i].SubItems[1], -1);
    if (p > -1) and (s > -1) then
    begin
      if p = OldPos then
        lvTXTFields.Items[i].SubItems[0] := IntToStr(NewPos)
      else if p + s = OldPos then
        lvTXTFields.Items[i].SubItems[1] := IntToStr(NewPos - p);
    end;
  end;
end;

procedure TQImport3WizardF.TXTViewerIntersectArrows(Sender: TObject;
  Position: Integer);
var
  i, p, s: Integer;
begin
  for i := 0 to lvTXTFields.Items.Count - 1 do
  begin
    p := StrToIntDef(lvTXTFields.Items[i].SubItems[0], -1);
    s := StrToIntDef(lvTXTFields.Items[i].SubItems[1], -1);
    if (p > -1) and (s > -1) and
       ((p = Position) or (p + s = Position)) then
    begin
      lvTXTFields.Items[i].SubItems[0] := EmptyStr;
      lvTXTFields.Items[i].SubItems[1] := EmptyStr;
    end;
  end;
end;

// CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV
// CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV
// CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV CSV

procedure TQImport3WizardF.CSVFillCombo;
var
  i: Integer;
begin
  if not (aiCSV in Wizard.AllowedImports) then Exit;
  if not QImportDestinationAssigned(False, ImportDestination, DataSet, DBGrid,
         ListView, StringGrid) then Exit;
  CSVClearCombo;
  for i := 0 to QImportDestinationColCount(False, ImportDestination, DataSet,
                  DBGrid, ListView, StringGrid) - 1 do
    with lvCSVFields.Items.Add do
    begin
      Caption := QImportDestinationColName(False, ImportDestination, DataSet,
        DBGrid, Self.ListView, StringGrid, GridCaptionRow, i);
      SubItems.Add(EmptyStr);
    end;
end;

procedure TQImport3WizardF.CSVClearCombo;
begin
  lvCSVFields.Items.Clear;
end;

procedure TQImport3WizardF.CSVTune;
begin
  ShowTip(tsCSVOptions, 6, 21, 39, tsCSVOptions.Width - 6, QImportLoadStr(QIW_CSV_Tip));

  pgImport.ActivePage := tsCSVOptions;
  laStep_07.Caption := Format(QImportLoadStr(QIW_Step), [1, 3]);

  if FNeedLoadFile then
    begin
      cmbCSVEncoding.ItemIndex := Integer(ctWinDefined);
      CSVFillGrid;
    end;

  if (lvCSVFields.Items.Count > 0) and
     not Assigned(lvCSVFields.Selected) then
  begin
    lvCSVFields.Items[0].Focused := True;
    lvCSVFields.Items[0].Selected := True;
  end;
  TuneButtons;
end;

function TQImport3WizardF.CSVReady: boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to lvCSVFields.Items.Count - 1 do
  begin
    Result := StrToIntDef(lvCSVFields.Items[i].SubItems[0], 0) > 0;
    if Result then Break;
  end;
end;

procedure TQImport3WizardF.CSVFillGrid;
var
{$IFDEF QI_UNICODE}
  F: TGpTextFile;
  TextBuffer: WideString;
  Strings: TqiStrings;
  lvCodePage: Integer;
{$ELSE}
  F: TextFile;
  Str: AnsiString;
  Strings: TStrings;
{$ENDIF}
  n, i, m: Integer;
  List: TList;
begin
  for i := 0 to sgrCSV.RowCount - 1 do
    for n := 0 to sgrCSV.ColCount - 1 do
      sgrCSV.Cells[n, i] := '';

  if not FileExists(FileName) then Exit;
{$IFDEF QI_UNICODE}
  F := TGpTextFile.CreateEx(FileName, FILE_ATTRIBUTE_NORMAL, GENERIC_READ, FILE_SHARE_READ);
  F.TryReadUpCRLF := True;
  F.Reset;
  lvCodePage:=SystemCTNames[FCSVEncoding];
  if lvCodePage <> -1 then
    F.Codepage:=lvCodePage;
  try
    Strings := TqiStringList.Create;
    try
      List := TList.Create;
      try
        n := 0;
        m := Wizard.CSVViewerRows;
        while not F.Eof and ((m <= 0) or (n <= m)) do
        begin
          TextBuffer := CsvReadLn(F, Comma, Quote);

          CSVStringToStrings(TextBuffer, Quote, Comma, Strings);
          for i := 0 to Strings.Count - 1 do
          begin
            if List.Count < i + 1 then
              List.Add(Pointer(Length(QImportLoadStr(QIW_CSV_GridCol))));
            if Integer(List[i]) < Length(Strings[i]) then
              List[i] := Pointer(Length(Strings[i]));
            if sgrCSV.ColCount < List.Count then
              sgrCSV.ColCount := List.Count;
            sgrCSV.Cells[i, n + 1] := Strings[i];
          end;
          Inc(n);
        end;
        sgrCSV.ColCount := List.Count;
        cbCSVColNumber.Items.Clear;
        cbCSVColNumber.Items.Add('');
        for i := 0 to List.Count - 1 do
        begin
          sgrCSV.ColWidths[i] := sgrCSV.Canvas.{$IFNDEF WIN64}TextWidthW{$ELSE}TextWidth{$ENDIF}('X') * (Integer(List[i]) + 1);
          sgrCSV.Cells[i, 0] := Format(QImportLoadStr(QIW_CSV_GridCol), [i + 1]);
          cbCSVColNumber.Items.Add(IntToStr(i + 1));
        end;
        sgrCSV.RowHeights[0] := 18;
        sgrCSV.RowCount := n + 1;
      finally
        List.Free;
      end;
    finally
      Strings.Free;
    end;
  finally
    F.Free;
  end;
{$ELSE}
  AssignFile(F, FileName);
  Reset(F);
  try
    Strings := TStringList.Create;
    try
      List := TList.Create;
      try
        n := 0;
        m := Wizard.CSVViewerRows;
        while not Eof(F) and ((m <= 0) or (n <= m)) do
        begin
          Readln(F, Str);
          CSVStringToStrings(Str, Quote, Comma, Strings);
          for i := 0 to Strings.Count - 1 do
          begin
            if List.Count < i + 1 then List.Add(Pointer(Length(QImportLoadStr(QIW_CSV_GridCol))));
            if Integer(List[i]) < Length(Strings[i]) then
              List[i] := Pointer(Length(Strings[i]));
            if sgrCSV.ColCount < List.Count then
              sgrCSV.ColCount := List.Count;
            sgrCSV.Cells[i, n + 1] := Strings[i];
          end;
          Inc(n);
        end;
        sgrCSV.ColCount := List.Count;
        cbCSVColNumber.Items.Clear;
        cbCSVColNumber.Items.Add(EmptyStr);
        for i := 0 to List.Count - 1 do
        begin
          sgrCSV.ColWidths[i] := sgrCSV.Canvas.TextWidth('X') * (Integer(List[i]) + 1);
          sgrCSV.Cells[i, 0] := Format(QImportLoadStr(QIW_CSV_GridCol), [i + 1]);
          cbCSVColNumber.Items.Add(IntToStr(i + 1));
        end;
        sgrCSV.RowHeights[0] := 18;
        sgrCSV.RowCount := n;
      finally
        List.Free;
      end;
    finally
      Strings.Free;
    end;
  finally
    CloseFile(F);
  end;
{$ENDIF}  
  FNeedLoadFile := False;
end;

function TQImport3WizardF.CSVCol: Integer;
begin
  Result := 0;
  if Assigned(lvCSVFields.Selected) then
    Result := StrToIntDef(lvCSVFields.Selected.SubItems[0], 0);
end;

procedure TQImport3WizardF.sgrCSVDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  X, Y: Integer;
begin
  X := Rect.Left + (Rect.Right - Rect.Left -
    sgrCSV.Canvas.TextWidth(sgrCSV.Cells[ACol, ARow])) div 2;
  Y := Rect.Top + (Rect.Bottom - Rect.Top -
    sgrCSV.Canvas.TextHeight(sgrCSV.Cells[ACol, ARow])) div 2;
  if gdFixed in State then
  begin
    if (ACol = CSVCol - 1) and (ARow = 0)
      then sgrCSV.Canvas.Font.Style := sgrCSV.Canvas.Font.Style + [fsBold]
      else sgrCSV.Canvas.Font.Style := sgrCSV.Canvas.Font.Style - [fsBold];
    sgrCSV.Canvas.FillRect(Rect);
    sgrCSV.Canvas.TextOut(X - 1, Y + 1, sgrCSV.Cells[ACol, ARow]);
  end
  else begin
    sgrCSV.DefaultDrawing := False;
    sgrCSV.Canvas.Brush.Color := clWindow;
    sgrCSV.Canvas.FillRect(Rect);
    sgrCSV.Canvas.Font.Color := clWindowText;
    sgrCSV.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2,
      sgrCSV.Cells[ACol, ARow]);

    if (ACol = CSVCol - 1) and (ARow > 0) then
    begin
      sgrCSV.Canvas.Font.Color := clHighLightText;
      sgrCSV.Canvas.Brush.Color := clHighLight;
      Rect.Bottom := Rect.Bottom + 1;
      sgrCSV.Canvas.FillRect(Rect);
      sgrCSV.Canvas.TextOut(Rect.Left + 2, Y, sgrCSV.Cells[ACol, ARow]);
      Rect.Bottom := Rect.Bottom - 1;
    end;
  end;
  if gdFocused in State
    then DrawFocusRect(sgrCSV.Canvas.Handle, Rect);
  sgrCSV.DefaultDrawing := True;
end;

procedure TQImport3WizardF.edtCSVSkipLinesChange(Sender: TObject);
begin
  CSVSkipLines := StrToIntDef(edtCSVSkipLines.Text, 0);
end;

procedure TQImport3WizardF.btnCSVClearClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvCSVFields.Items.Count - 1 do
    lvCSVFields.Items[i].SubItems[0] := EmptyStr;
  lvCSVFieldsChange(lvCSVFields, lvCSVFields.Selected, ctState);
end;

procedure TQImport3WizardF.sgrCSVMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: Integer;
begin
  sgrCSV.MouseToCell(X, Y, ACol, ARow);
  if not Assigned(lvCSVFields.Selected) then Exit;
  if CSVCol = ACol + 1 then
    lvCSVFields.Selected.SubItems[0] := EmptyStr
  else
    lvCSVFields.Selected.SubItems[0] := IntToStr(ACol + 1);
  lvCSVFieldsChange(lvCSVFields, lvCSVFields.Selected, ctState);
end;

procedure TQImport3WizardF.lvCSVFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;
  cbCSVColNumber.OnChange := nil;
  try
    if Item.SubItems.Count > 0 then
      cbCSVColNumber.ItemIndex := cbCSVColNumber.Items.IndexOf(Item.SubItems[0]);
    sgrCSV.Repaint;
    TuneButtons;
  finally
    cbCSVColNumber.OnChange := cbCSVColNumberChange;;
  end;
end;

procedure TQImport3WizardF.cbCSVColNumberChange(Sender: TObject);
begin
  if not Assigned(lvCSVFields.Selected) then Exit;
  lvCSVFields.Selected.SubItems[0] := cbCSVColNumber.Text;
  sgrCSV.Repaint;
end;

procedure TQImport3WizardF.btnCSVAutoFillClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvCSVFields.Items.Count - 1 do
    if (i <= sgrCSV.ColCount - 1)
      then lvCSVFields.Items[i].SubItems[0] := IntToStr(i + 1)
      else lvCSVFields.Items[i].SubItems[0] := EmptyStr;
  lvCSVFieldsChange(lvCSVFields, lvCSVFields.Selected, ctState);
end;

procedure TQImport3WizardF.SetCSVSkipLines(const Value: Integer);
begin
  if FCSVSkipLines <> Value then
  begin
    FCSVSkipLines := Value;
    edtCSVSkipLines.Text := IntToStr(FCSVSkipLines);
  end;
end;

// DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF
// DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF
// DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF DBF

procedure TQImport3WizardF.DBFFillList;
var
  i: Integer;
begin
  if not (aiDBF in Wizard.AllowedImports) then Exit;
  DBFClearList;
  for i := 0 to QImportDestinationColCount(False, ImportDestination, DataSet,
                  DBGrid, ListView, StringGrid) - 1 do
    lstDBFDataSet.Items.Add.Caption :=
      QImportDestinationColName(False, ImportDestination, DataSet, DBGrid,
        ListView, StringGrid, GridCaptionRow, i);
  if lstDBFDataSet.Items.Count > 0 then
  begin
    lstDBFDataSet.Items[0].Focused := True;
    lstDBFDataSet.Items[0].Selected := True;
  end;
end;

procedure TQImport3WizardF.DBFClearList;
var
  i: Integer;
begin
  for i := lstDBFDataSet.Items.Count - 1 downto 0 do
    lstDBFDataSet.Items.Delete(i);
end;

procedure TQImport3WizardF.DBFFillTableList;
var
  DBF: TDBFRead;
  i: Integer;
begin
  DBFClearTableList;
  DBF := TDBFRead.Create(FileName);
  try
    for i := 0 to DBF.FieldCount - 1 do
      lstDBF.Items.Add.Caption := string(DBF.FieldName[i]);
    if lstDBF.Items.Count > 0 then
    begin
      lstDBF.Items[0].Focused := True;;
      lstDBF.Items[0].Selected := True;;
    end;
    lstDBF.Columns[0].Caption := ExtractFileName(FileName);
  finally
    DBF.Free;
  end;
  FNeedLoadFile := False;
end;

procedure TQImport3WizardF.DBFClearTableList;
var
 i: Integer;
begin
  for i := lstDBF.Items.Count - 1 downto 0 do
    lstDBF.Items.Delete(i);
end;

procedure TQImport3WizardF.DBFTune;
begin
  ShowTip(tsDBFOptions, 6, 21, 39, tsDBFOptions.Width - 6, QImportLoadStr(QIW_DBF_Tip));

  if FNeedLoadFile then DBFFillTableList;
  pgImport.ActivePage := tsDBFOptions;
  laStep_03.Caption := Format(QImportLoadStr(QIW_Step), [1,3]);
  TuneButtons;
end;

procedure TQImport3WizardF.lstDBFDataSetChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  TuneButtons;
end;

procedure TQImport3WizardF.lstDBFChange(Sender: TObject; Item: TListItem;
  Change: TItemChange);
begin
  TuneButtons;
end;

procedure TQImport3WizardF.lstDBFMapChange(Sender: TObject; Item: TListItem;
  Change: TItemChange);
begin
  TuneButtons;
end;

procedure TQImport3WizardF.SetDBFSkipDeleted(const Value: Boolean);
begin
  if FDBFSkipDeleted <> Value then begin
    FDBFSkipDeleted := Value;
    chDBFSkipDeleted.Checked := FDBFSkipDeleted;
  end;
end;

procedure TQImport3WizardF.chDBFSkipDeletedClick(Sender: TObject);
begin
  DBFSkipDeleted := chDBFSkipDeleted.Checked;
end;

procedure TQImport3WizardF.pbDBFAddPaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: Integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := True;
    if bDBFAdd.Enabled
      then i := 4
      else i := 12;
    ilWizard.GetBitmap(i, Bmp);
    pbDBFAdd.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TQImport3WizardF.bDBFAddMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbDBFAdd.Left := pbDBFAdd.Left + 1;
  pbDBFAdd.Top := pbDBFAdd.Top + 1;
end;

procedure TQImport3WizardF.bDBFAddMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbDBFAdd.Left := pbDBFAdd.Left - 1;
  pbDBFAdd.Top := pbDBFAdd.Top - 1;
end;

procedure TQImport3WizardF.pbDBFAutoFillPaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: Integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := True;
    if bDBFAutoFill.Enabled
      then i := 9
      else i := 17;
    ilWizard.GetBitmap(i, Bmp);
    pbDBFAutoFill.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TQImport3WizardF.bDBFAutoFillMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbDBFAutoFill.Left := pbDBFAutoFill.Left + 1;
  pbDBFAutoFill.Top := pbDBFAutoFill.Top + 1;
end;

procedure TQImport3WizardF.bDBFAutoFillMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbDBFAutoFill.Left := pbDBFAutoFill.Left - 1;
  pbDBFAutoFill.Top := pbDBFAutoFill.Top - 1;
end;

procedure TQImport3WizardF.pbDBFRemovePaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: Integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := True;
    if bDBFRemove.Enabled
      then i := 6
      else i := 14;
    ilWizard.GetBitmap(i, Bmp);
    pbDBFRemove.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TQImport3WizardF.bDBFRemoveMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbDBFRemove.Left := pbDBFRemove.Left + 1;
  pbDBFRemove.Top := pbDBFRemove.Top + 1;
end;

procedure TQImport3WizardF.bDBFRemoveMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbDBFRemove.Left := pbDBFRemove.Left - 1;
  pbDBFRemove.Top := pbDBFRemove.Top - 1;
end;

procedure TQImport3WizardF.pbDBFClearPaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: Integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := True;
    if bDBFClear.Enabled
      then i := 7
      else i := 15;
    ilWizard.GetBitmap(i, Bmp);
    pbDBFClear.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TQImport3WizardF.bDBFClearMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbDBFClear.Left := pbDBFClear.Left + 1;
  pbDBFClear.Top := pbDBFClear.Top + 1;
end;

procedure TQImport3WizardF.bDBFClearMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbDBFClear.Left := pbDBFClear.Left - 1;
  pbDBFClear.Top := pbDBFClear.Top - 1;
end;

// XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML
// XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML
// XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML XML
procedure TQImport3WizardF.SetXMLWriteOnFly(const Value: boolean);
begin
  if FXMLWriteOnFly <> Value then begin
    FXMLWriteOnFly := Value;
    chXMLWriteOnFly.Checked := FXMLWriteOnFly;
  end;
end;

procedure TQImport3WizardF.chXMLWriteOnFlyClick(Sender: TObject);
begin
  XMLWriteOnFly := chXMLWriteOnFly.Checked;
end;

procedure TQImport3WizardF.bXMLAddClick(Sender: TObject);
begin
  with lvXMLMap.Items.Add do begin
    Caption := lvXMLDataSet.Selected.Caption;
    SubItems.Add('=');
    SubItems.Add(lvXML.Selected.Caption);
    Focused := True;
    Selected := True;
  end;
  lvXMLDataSet.Items.Delete(lvXMLDataSet.Selected.Index);
  if lvXMLDataSet.Items.Count > 0 then begin
    lvXMLDataSet.Items[0].Focused := True;
    lvXMLDataSet.Items[0].Selected := True;
  end;
  if lvXMLMap.Items.Count > 0 then begin
    lvXMLMap.Items[0].Focused := True;
    lvXMLMap.Items[0].Selected := True;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.bXMLAutoFillClick(Sender: TObject);
var
  i, N: Integer;
begin
  lvXMLDataSet.Items.BeginUpdate;
  try
    lvXML.Items.BeginUpdate;
    try
      lvXMLMap.Items.BeginUpdate;
      try
        lvXMLMap.Items.Clear;
        XMLClearList;
        XMLFillList;
        N := lvXML.Items.Count;
        if N > lvXMLDataSet.Items.Count
          then N := lvXMLDataSet.Items.Count;
        for i := N - 1 downto 0 do begin
          with lvXMLMap.Items.Insert(0) do begin
            Caption := lvXMLDataSet.Items[i].Caption;
            SubItems.Add('=');
            SubItems.Add(lvXML.Items[i].Caption);
          end;
          lvXMLDataSet.Items[i].Delete;
        end;
        if lvXMLMap.Items.Count > 0 then begin
          lvXMLMap.Items[0].Focused := True;
          lvXMLMap.Items[0].Selected := True;
        end;
      finally
        lvXMLMap.Items.EndUpdate;
      end;
    finally
      lvXML.Items.EndUpdate;
    end;
  finally
    lvXMLDataSet.Items.EndUpdate;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.bXMLRemoveClick(Sender: TObject);
begin
  lvXMLDataSet.Items.Add.Caption := lvXMLMap.Selected.Caption;
  lvXMLMap.Items.Delete(lvXMLMap.Selected.Index);
  if lvXMLMap.Items.Count > 0 then begin
    lvXMLMap.Items[0].Focused := True;
    lvXMLMap.Items[0].Selected := True;
  end;
  if lvXMLDataSet.Items.Count > 0 then begin
    lvXMLDataSet.Items[0].Focused := True;
    lvXMLDataSet.Items[0].Selected := True;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.bXMLClearClick(Sender: TObject);
begin
  lvXMLDataSet.Items.BeginUpdate;
  try
    lvXMLMap.Items.BeginUpdate;
    try
      lvXMLMap.Items.Clear;
      XMLClearList;
      XMLFillList;
    finally
      lvXMLMap.Items.EndUpdate;
    end;
  finally
    lvXMLDataSet.Items.EndUpdate;
  end;
  TuneButtons;
end;

function TQImport3WizardF.XMLReady: boolean;
begin
  Result := lvXMLMap.Items.Count > 0;
end;

procedure TQImport3WizardF.XMLFillList;
var
  i: Integer;
begin
  if not (aiXML in Wizard.AllowedImports) then Exit;
  XMLClearList;
  for i := 0 to QImportDestinationColCount(False, ImportDestination, DataSet,
                  DBGrid, ListView, StringGrid) - 1 do
    lvXMLDataSet.Items.Add.Caption :=
      QImportDestinationColName(False, ImportDestination, DataSet, DBGrid,
        ListView, StringGrid, GridCaptionRow, i);
  if lvXMLDataSet.Items.Count > 0 then  begin
    lvXMLDataSet.Items[0].Focused := True;
    lvXMLDataSet.Items[0].Selected := True;
  end;
end;

procedure TQImport3WizardF.XMLClearList;
var
  i: Integer;
begin
  for i := lvXMLDataSet.Items.Count - 1 downto 0 do
    lvXMLDataSet.Items.Delete(i);
end;

procedure TQImport3WizardF.XMLFillTableList;
var
  XML: TXMLFile;
  i: Integer;
  str: string;
begin
  XMLClearTableList;
  XML := TXMLFile.Create;
  try
    XML.FileName := FileName;
    XML.Clear;
    XML.FileType := TQIXMLDocType(cbXMLDocumentType.ItemIndex);
    XML.Load(True);
    if not Assigned(XML.Fields) then
      raise EQImportError.Create(QImportLoadStr(QIE_XMLFieldListEmpty));
       

    for i := 0 to XML.FieldCount - 1 do begin
      str := EmptyStr;
      if XML.Fields[i].Attributes.IndexOfName('attrname') > -1 then
        str := 'attrname'
      else if XML.Fields[i].Attributes.IndexOfName('FieldName') > -1 then
        str := 'FieldName';
      if str = EmptyStr then raise EQImportError.Create('XML read with errors');

      lvXML.Items.Add.Caption := XML.Fields[i].Attributes.Values[str];
    end;
    if lvXML.Items.Count > 0 then begin
      lvXML.Items[0].Focused := True;;
      lvXML.Items[0].Selected := True;;
    end;
    lvXML.Columns[0].Caption := ExtractFileName(FileName);
  finally
    XML.Free;
  end;
  FNeedLoadFile := False;
end;

procedure TQImport3WizardF.XMLClearTableList;
var
 i: Integer;
begin
  for i := lvXML.Items.Count - 1 downto 0 do
    lvXML.Items.Delete(i);
end;

procedure TQImport3WizardF.XMLTune;
begin
  ShowTip(tsXMLOptions, 6, 21, 39, tsXMLOptions.Width - 6, QImportLoadStr(QIW_XML_Tip));

  if FNeedLoadFile then XMLFillTableList;
  pgImport.ActivePage := tsXMLOptions;
  laStep_05.Caption := Format(QImportLoadStr(QIW_Step), [1,3]);
  TuneButtons;
end;

procedure TQImport3WizardF.pbXMLAddPaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: Integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := True;
    if bXMLAdd.Enabled
      then i := 4
      else i := 12;
    ilWizard.GetBitmap(i, Bmp);
    pbXMLAdd.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TQImport3WizardF.bXMLAddMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbXMLAdd.Left := pbXMLAdd.Left + 1;
  pbXMLAdd.Top := pbXMLAdd.Top + 1;
end;

procedure TQImport3WizardF.bXMLAddMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbXMLAdd.Left := pbXMLAdd.Left - 1;
  pbXMLAdd.Top := pbXMLAdd.Top - 1;
end;

procedure TQImport3WizardF.pbXMLAutoFillPaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: Integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := True;
    if bXMLAutoFill.Enabled
      then i := 9
      else i := 17;
    ilWizard.GetBitmap(i, Bmp);
    pbXMLAutoFill.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TQImport3WizardF.bXMLAutoFillMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbXMLAutoFill.Left := pbXMLAutoFill.Left + 1;
  pbXMLAutoFill.Top := pbXMLAutoFill.Top + 1;
end;

procedure TQImport3WizardF.bXMLAutoFillMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbXMLAutoFill.Left := pbXMLAutoFill.Left - 1;
  pbXMLAutoFill.Top := pbXMLAutoFill.Top - 1;
end;

procedure TQImport3WizardF.pbXMLRemovePaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: Integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := True;
    if bXMLRemove.Enabled
      then i := 6
      else i := 14;
    ilWizard.GetBitmap(i, Bmp);
    pbXMLRemove.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TQImport3WizardF.bXMLRemoveMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbXMLRemove.Left := pbXMLRemove.Left + 1;
  pbXMLRemove.Top := pbXMLRemove.Top + 1;
end;

procedure TQImport3WizardF.bXMLRemoveMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbXMLRemove.Left := pbXMLRemove.Left - 1;
  pbXMLRemove.Top := pbXMLRemove.Top - 1;
end;

procedure TQImport3WizardF.pbXMLClearPaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: Integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := True;
    if bXMLClear.Enabled
      then i := 7
      else i := 15;
    ilWizard.GetBitmap(i, Bmp);
    pbXMLClear.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TQImport3WizardF.bXMLClearMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbXMLClear.Left := pbXMLClear.Left + 1;
  pbXMLClear.Top := pbXMLClear.Top + 1;
end;

procedure TQImport3WizardF.bXMLClearMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbXMLClear.Left := pbXMLClear.Left - 1;
  pbXMLClear.Top := pbXMLClear.Top - 1;
end;

procedure TQImport3WizardF.tbtAddReplacementClick(Sender: TObject);
var
  TextToFind, ReplaceWith: String;
  IgnoreCase: boolean;
  R: TQImportReplacement;
begin
  TextToFind := EmptyStr;
  ReplaceWith := EmptyStr;
  IgnoreCase := False;
  if ReplacementEdit(TextToFind, ReplaceWith, IgnoreCase) then begin
    R := TQImportFieldFormat(lstFormatFields.ItemFocused.Data).Replacements.Add;
    R.TextToFind := TextToFind;
    R.ReplaceWith := ReplaceWith;
    R.IgnoreCase := IgnoreCase;
    with lvReplacements.Items.Add do begin
      Caption := R.TextToFind;
      SubItems.Add(R.ReplaceWith);
      if R.IgnoreCase
        then SubItems.Add(QImportLoadStr(QIWDF_IgnoreCase_Yes))
        else SubItems.Add(QImportLoadStr(QIWDF_IgnoreCase_No));
      ImageIndex := 3;
      Data := R;
      Focused := True;
      Selected := True;
    end;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.lvReplacementsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  TuneButtons;
end;

procedure TQImport3WizardF.tbtEditReplacementClick(Sender: TObject);
var
  TextToFind, ReplaceWith: String;
  IgnoreCase: boolean;
  R: TQImportReplacement;
begin
  R := TQImportReplacement(lvReplacements.ItemFocused.Data);
  TextToFind := R.TextToFind;
  ReplaceWith := R.ReplaceWith;
  IgnoreCase := R.IgnoreCase;
  if ReplacementEdit(TextToFind, ReplaceWith, IgnoreCase) then begin
    R.TextToFind := TextToFind;
    R.ReplaceWith := ReplaceWith;
    R.IgnoreCase := IgnoreCase;
    with lvReplacements.ItemFocused do begin
      Caption := R.TextToFind;
      SubItems[0] := R.ReplaceWith;
      if R.IgnoreCase
        then SubItems[1] := QImportLoadStr(QIWDF_IgnoreCase_Yes)
        else SubItems[1] := QImportLoadStr(QIWDF_IgnoreCase_No);
    end;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.tbtDelReplacementClick(Sender: TObject);
begin
  TQImportReplacement(lvReplacements.ItemFocused.Data).Free;
  lvReplacements.ItemFocused.Delete;
  TuneButtons;
end;

procedure TQImport3WizardF.lvReplacementsDblClick(Sender: TObject);
begin
  if tbtEditReplacement.Enabled then
    tbtEditReplacement.Click; 
end;

procedure TQImport3WizardF.lvXLSFieldsEnter(Sender: TObject);
begin
  XLSRepaintCurrentGrid;
end;

procedure TQImport3WizardF.lvXLSFieldsExit(Sender: TObject);
begin
  XLSRepaintCurrentGrid;
end;

procedure TQImport3WizardF.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if Wizard.ConfirmOnCancel then
  begin
    CanClose := (Application.MessageBox(PChar(QImportLoadStr(QIW_CancelWizard)),
                                        PChar(QImportLoadStr(QIW_Question)),
          MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2) = ID_YES);
  end;
  if CanClose and AutoSaveTemplate then
    SaveTemplateToFile(TemplateFileName);
end;

function TQImport3WizardF.AllowedImportFileType(
  const AFileName: qiString): Boolean;
var
  ext: qiString;
begin
  ext := LowerCase(ExtractFileExt(AFileName));
  Result := (ext = XLS_EXT) or
            (ext = XLSX_EXT) or
            (ext = DOCX_EXT) or
            (ext = ODS_EXT) or
            (ext = ODT_EXT) or
            (ext = DBF_EXT) or
            (ext = DB_EXT) or
            (ext = HTM_EXT) or
            (ext = HTML_EXT) or 
            (ext = XML_EXT) or
            (ext = TXT_EXT) or
            (ext = CSV_EXT) or
            (ext = MDB_EXT);
end;

function TQImport3WizardF.ImportTypeEquivFileType(const AFileName: qiString): Boolean;
var
  ext: qiString;
begin
  ext := LowerCase(ExtractFileExt(AFileName));
  CheckExtension(ext);
  Result := ImportTypeStr(ImportType) = ext;
end;

function TQImport3WizardF.ImportTypeStr(
  AImportType: TAllowedImport): string;
begin
  Result := '';
  case AImportType of
    aiXLS: Result := XLS_EXT;
    aiXlsx: Result := XLSX_EXT;
    aiDocx: Result := DOCX_EXT;
    aiODS: Result := ODS_EXT;
    aiODT: Result := ODT_EXT;
    aiDBF: Result := DBF_EXT;
    aiHTML: Result := HTML_EXT;
    aiXML: Result := XML_EXT;
    aiXMLDoc: Result := XML_EXT;
    aiTXT: Result := TXT_EXT;
    aiCSV: Result := CSV_EXT;
    aiAccess: Result := MDB_EXT;
  end;
end;

procedure TQImport3WizardF.CheckExtension(var Ext: qiString);
begin
  if Ext = '.htm' then Ext := Ext + 'l';
end;

{$IFDEF HTML}
{$IFDEF VCL6}
// HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML
// HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML
// HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML HTML

procedure TQImport3WizardF.lvHTMLFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;
  cbHTMLColNumber.OnChange := nil;
  try
    if Item.SubItems.Count > 0 then
      cbHTMLColNumber.ItemIndex := cbHTMLColNumber.Items.IndexOf(Item.SubItems[0]);
    sgrHTML.Repaint;
    TuneButtons;
  finally
    cbHTMLColNumber.OnChange := cbHTMLColNumberChange;
  end;
end;

procedure TQImport3WizardF.tbtHTMLAutoFillClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvHTMLFields.Items.Count - 1 do
    if (i <= sgrHTML.ColCount - 1) then
      lvHTMLFields.Items[i].SubItems[0] := IntToStr(i + 1)
    else
      lvHTMLFields.Items[i].SubItems[0] := EmptyStr;
  lvHTMLFieldsChange(lvHTMLFields, lvHTMLFields.Selected, ctState);
end;

procedure TQImport3WizardF.tbtHTMLClearClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvHTMLFields.Items.Count - 1 do
    lvHTMLFields.Items[i].SubItems[0] := EmptyStr;
  lvHTMLFieldsChange(lvHTMLFields, lvHTMLFields.Selected, ctState);
end;

procedure TQImport3WizardF.cbHTMLTableNumberChange(Sender: TObject);
begin
  HTMLFillGrid;
end;

procedure TQImport3WizardF.cbHTMLColNumberChange(Sender: TObject);
begin
  if not Assigned(lvHTMLFields.Selected) then Exit;
  lvHTMLFields.Selected.SubItems[0] := cbHTMLColNumber.Text;
  sgrHTML.Repaint;
end;

procedure TQImport3WizardF.edtHTMLSkipLinesChange(Sender: TObject);
begin
  HTMLSkipLines := StrToIntDef(edtHTMLSkipLines.Text, 0);
end;

procedure TQImport3WizardF.sgrHTMLDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  X, Y: Integer;
begin
  X := Rect.Left + (Rect.Right - Rect.Left -
   sgrHTML.Canvas.TextWidth(sgrHTML.Cells[ACol, ARow])) div 2;
  Y := Rect.Top + (Rect.Bottom - Rect.Top -
    sgrHTML.Canvas.TextHeight(sgrHTML.Cells[ACol, ARow])) div 2;
  if gdFixed in State then
  begin
    sgrHTML.Canvas.FillRect(Rect);
    sgrHTML.Canvas.TextOut(X - 1, Y + 1, sgrHTML.Cells[ACol, ARow]);
  end
  else begin
    sgrHTML.DefaultDrawing := False;
    sgrHTML.Canvas.Brush.Color := clWindow;
    sgrHTML.Canvas.FillRect(Rect);
    sgrHTML.Canvas.Font.Color := clWindowText;
    sgrHTML.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2, sgrHTML.Cells[ACol, ARow]);

    if (ACol = HTMLCol - 1) then
    begin
      sgrHTML.Canvas.Font.Color := clHighLightText;
      sgrHTML.Canvas.Brush.Color := clHighLight;
      Rect.Bottom := Rect.Bottom + 1;
      sgrHTML.Canvas.FillRect(Rect);
      sgrHTML.Canvas.TextOut(Rect.Left + 2, Y, sgrHTML.Cells[ACol, ARow]);
      Rect.Bottom := Rect.Bottom - 1;
    end;
  end;
  if gdFocused in State then
    DrawFocusRect(sgrHTML.Canvas.Handle, Rect);
  sgrHTML.DefaultDrawing := True;
end;

procedure TQImport3WizardF.sgrHTMLMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: Integer;
begin
  sgrHTML.MouseToCell(X, Y, ACol, ARow);
  if not Assigned(lvHTMLFields.Selected) then Exit;
  if HTMLCol = ACol + 1 then
    lvHTMLFields.Selected.SubItems[0] := EmptyStr
  else
    if ACol > -1 then
      lvHTMLFields.Selected.SubItems[0] := IntToStr(ACol + 1);

  lvHTMLFieldsChange(lvHTMLFields, lvHTMLFields.Selected, ctState);
end;

procedure TQImport3WizardF.HTMLFillList;
var
  i: Integer;
begin
  if not (aiHTML in Wizard.AllowedImports) then Exit;
  HTMLClearList;
  for i := 0 to QImportDestinationColCount(False, ImportDestination, DataSet,
                  DBGrid, ListView, StringGrid) - 1 do
    with lvHTMLFields.Items.Add do
    begin
      Caption := QImportDestinationColName(False, ImportDestination, DataSet,
        DBGrid, Self.ListView, StringGrid, GridCaptionRow, i);
      SubItems.Add(EmptyStr);
    end;
    
  if lvHTMLFields.Items.Count > 0 then  begin
    lvHTMLFields.Items[0].Focused := True;
    lvHTMLFields.Items[0].Selected := True;
  end;
end;

procedure TQImport3WizardF.HTMLClearList;
var
  i: Integer;
begin
  for i := lvHTMLFields.Items.Count - 1 downto 0 do
    lvHTMLFields.Items.Delete(i);
end;

procedure TQImport3WizardF.HTMLTune;
begin
  ShowTip(tsHTMLOptions, 6, 21, 39, tsHTMLOptions.Width - 6, QImportLoadStr(QIW_HTML_Tip));
  pgImport.ActivePage := tsHTMLOptions;
  laStep_08.Caption := Format(QImportLoadStr(QIW_Step), [1, 3]);

  if FNeedLoadFile then HTMLFillGrid;

  if (lvHTMLFields.Items.Count > 0) and
     not Assigned(lvHTMLFields.Selected) then
  begin
    lvHTMLFields.Items[0].Focused := True;
    lvHTMLFields.Items[0].Selected := True;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.HTMLFillGrid;
var
  F: TForm;
  Start, Finish: TDateTime;
  i, j:  Integer;
begin
  if not FileExists(FileName) then Exit;

  if FNeedLoadFile then
  begin
    Start := Now;
    F := ShowLoading(Self, FFileName);
    try
      Application.ProcessMessages;
      FHTMLFile.Clear;
      FHTMLFile.FileName := FileName;
      FHTMLFile.Load(0);
      cbHTMLTableNumber.Items.Clear;
      if FHTMLFile.TableList.Count >= 0 then
      begin
        for i := 0 to FHTMLFile.TableList.Count - 1 do
          cbHTMLTableNumber.Items.Add(IntToStr(Succ(i)));
        cbHTMLTableNumber.ItemIndex := 0;
      end;
    finally
      Finish := Now;
      while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
        Finish := Now;
      if Assigned(F) then
        F.Free;
    end;
  end;

  if FHTMLFile.TableList.Count >= 0 then
  begin
    sgrHTML.ColCount := 1;
    sgrHTML.RowCount := Min(FHTMLFile.TableList[cbHTMLTableNumber.ItemIndex].Rows.Count, 30);
    for i := 0 to sgrHTML.RowCount - 1 do
    begin
      if sgrHTML.ColCount < FHTMLFile.TableList[cbHTMLTableNumber.ItemIndex].Rows[i].Cells.Count then
        sgrHTML.ColCount := FHTMLFile.TableList[cbHTMLTableNumber.ItemIndex].Rows[i].Cells.Count;
      for j := 0 to FHTMLFile.TableList[cbHTMLTableNumber.ItemIndex].Rows[i].Cells.Count - 1 do
        sgrHTML.Cells[j, i] := FHTMLFile.TableList[cbHTMLTableNumber.ItemIndex].Rows[i].Cells[j].Text;
    end;
    HTMLFillComboColumn;
  end;
  FNeedLoadFile := False;
end;

procedure TQImport3WizardF.HTMLFillComboColumn;
var
  i: Integer;
begin
  if not (aiHTML in Wizard.AllowedImports) then Exit;
  cbHTMLColNumber.Clear;
  cbHTMLColNumber.Items.Add('');
  cbHTMLColNumber.ItemIndex := 0;
  for i := 0 to sgrHTML.ColCount - 1 do
    cbHTMLColNumber.Items.Add(IntToStr(Succ(i)));
end;

procedure TQImport3WizardF.SetHTMLSkipLines(const Value: Integer);
begin
  if FHTMLSkipLines <> Value then
  begin
    FHTMLSkipLines := Value;
    edtHTMLSkipLines.Text := IntToStr(FHTMLSkipLines);
  end;
end;

function TQImport3WizardF.HTMLReady: boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to lvHTMLFields.Items.Count - 1 do
  begin
    Result := StrToIntDef(lvHTMLFields.Items[i].SubItems[0], 0) > 0;
    if Result then Break;
  end;
end;

function TQImport3WizardF.HTMLCol: Integer;
begin
  Result := 0;
  if Assigned(lvHTMLFields.Selected) then
    Result := StrToIntDef(lvHTMLFields.Selected.SubItems[0], 0);
end;
{$ENDIF}
{$ENDIF}

{$IFDEF XMLDOC}
{$IFDEF VCL6}
// XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc
// XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc
// XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc XMLDoc

procedure TQImport3WizardF.XMLDocTune;
begin
  ShowTip(tsXMLDocOptions, 6, 21, 39, tsXMLDocOptions.Width - 6, QImportLoadStr(QIW_XMLDoc_Tip));
  pgImport.ActivePage := tsXMLDocOptions;
  laStep_10.Caption := Format(QImportLoadStr(QIW_Step), [1, 3]);

  if FNeedLoadFile then XMLDocLoadFile;

  if (lvXMLDocFields.Items.Count > 0) and
     not Assigned(lvXMLDocFields.Selected) then
  begin
    lvXMLDocFields.Items[0].Focused := True;
    lvXMLDocFields.Items[0].Selected := True;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.XMLDocLoadFile;
var
  F: TForm;
  Start, Finish: TDateTime;
begin
  FXMLDocFile.FileName := FFileName;
  Start := Now;
  F := ShowLoading(Self, FFileName);
  try
    Application.ProcessMessages;
    FXMLDocFile.Load;
    FNeedLoadFile := False;
    XMLDocClearGrid;
  finally
    Finish := Now;
    while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
      Finish := Now;
    F.Free;
  end;
end;

procedure TQImport3WizardF.XMLDocFillList;
var
  i: Integer;
begin
  if not (aiXMLDoc in Wizard.AllowedImports) then Exit;
  XMLDocClearList;
  for i := 0 to QImportDestinationColCount(False, ImportDestination, DataSet,
                  DBGrid, ListView, StringGrid) - 1 do
    with lvXMLDocFields.Items.Add do
    begin
      Caption := QImportDestinationColName(False, ImportDestination, DataSet,
        DBGrid, Self.ListView, StringGrid, GridCaptionRow, i);
      SubItems.Add(EmptyStr);
    end;
    
  if lvXMLDocFields.Items.Count > 0 then
  begin
    lvXMLDocFields.Items[0].Focused := True;
    lvXMLDocFields.Items[0].Selected := True;
  end;
end;

procedure TQImport3WizardF.XMLDocClearList;
var
  i: Integer;
begin
  for i := lvXMLDocFields.Items.Count - 1 downto 0 do
    lvXMLDocFields.Items.Delete(i);
end;

procedure TQImport3WizardF.XMLDocFillGrid;
begin
  XMLDocClearGrid;
  FillStringGrid(sgrXMLDoc, FXMLDocFile);
  XMLDocFillComboColumn;
end;

procedure TQImport3WizardF.XMLDocClearGrid;
var
  i: Integer;
begin
  for i := 0 to sgrXMLDoc.ColCount - 1 do
    sgrXMLDoc.Cols[i].Clear;
  sgrXMLDoc.ColCount := 2;
  sgrXMLDoc.FixedCols := 1;
  sgrXMLDoc.RowCount := 2;
  sgrXMLDoc.FixedRows := 1;
end;

procedure TQImport3WizardF.XMLDocFillComboColumn;
var
  i: Integer;
begin
  cbXMLDocColNumber.Clear;
  cbXMLDocColNumber.Items.Add('');
  cbXMLDocColNumber.ItemIndex := 0;
  for i := 0 to sgrXMLDoc.ColCount - 1 do
    cbXMLDocColNumber.Items.Add(IntToStr(Succ(i)));
end;

function TQImport3WizardF.XMLDocReady: boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to lvXMLDocFields.Items.Count - 1 do
  begin
    Result := StrToIntDef(lvXMLDocFields.Items[i].SubItems[0], 0) > 0;
    if Result then Break;
  end;
end;

function TQImport3WizardF.XMLDocCol: Integer;
begin
  Result := 0;
  if Assigned(lvXMLDocFields.Selected) then
    Result := StrToIntDef(lvXMLDocFields.Selected.SubItems[0], 0);
end;

procedure TQImport3WizardF.SetXMLDocSkipLines(const Value: Integer);
begin
  if FXMLDocSkipLines <> Value then
  begin
    FXMLDocSkipLines := Value;
    edXMLDocSkipLines.Text := IntToStr(FXMLDocSkipLines);
  end;
end;

procedure TQImport3WizardF.SetXMLDocXPath(const Value: qiString);
var
  F: TForm;
  Start, Finish: TDateTime;
begin
  if Value <> EmptyStr then
  begin
    FXMLDocXPath := Value;
    edXMLDocXPath.Text := Value;
    Start := Now;
    F := ShowLoading(Self, Value);
    try
      Application.ProcessMessages;
      FXMLDocFile.XPath := Value;
     finally
      Finish := Now;
      while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
        Finish := Now;
      F.Free;
    end;
    XMLDocFillGrid;
  end;
end;

procedure TQImport3WizardF.SetXMLDocDataLocation(
  const Value: TXMLDataLocation);
begin
  if Value <> FXMLDocDataLocation then
  begin
    FXMLDocDataLocation := Value;
    cbXMLDocDataLocation.ItemIndex := Integer(Value);
    FXMLDocFile.DataLocation := Value;
  end;
end;

procedure TQImport3WizardF.edXMLDocSkipLinesChange(Sender: TObject);
begin
  XMLDocSkipLines := StrToIntDef(edXMLDocSkipLines.Text, 0);
end;

procedure TQImport3WizardF.cbXMLDocColNumberChange(Sender: TObject);
begin
  if not Assigned(lvXMLDocFields.Selected) then Exit;
  lvXMLDocFields.Selected.SubItems[0] := cbXMLDocColNumber.Text;
  sgrXMLDoc.Repaint;
end;

procedure TQImport3WizardF.sgrXMLDocDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  X, Y: integer;
begin
  X := Rect.Left + (Rect.Right - Rect.Left -
   sgrXMLDoc.Canvas.TextWidth(sgrXMLDoc.Cells[ACol, ARow])) div 2;
  Y := Rect.Top + (Rect.Bottom - Rect.Top -
    sgrXMLDoc.Canvas.TextHeight(sgrXMLDoc.Cells[ACol, ARow])) div 2;
  if gdFixed in State then
  begin
    if (ACol = XMLDocCol - 1) and (ARow = 0)
      then sgrXMLDoc.Canvas.Font.Style := sgrXMLDoc.Canvas.Font.Style + [fsBold]
      else sgrXMLDoc.Canvas.Font.Style := sgrXMLDoc.Canvas.Font.Style - [fsBold];
    sgrXMLDoc.Canvas.FillRect(Rect);
    sgrXMLDoc.Canvas.TextOut(X - 1, Y + 1, sgrXMLDoc.Cells[ACol, ARow]);
  end
  else begin
    sgrXMLDoc.DefaultDrawing := false;
    sgrXMLDoc.Canvas.Brush.Color := clWindow;
    sgrXMLDoc.Canvas.FillRect(Rect);
    sgrXMLDoc.Canvas.Font.Color := clWindowText;
    sgrXMLDoc.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2,
      sgrXMLDoc.Cells[ACol, ARow]);

    if (ACol = XMLDocCol - 1) and (ARow > 0) then
    begin
      sgrXMLDoc.Canvas.Font.Color := clHighLightText;
      sgrXMLDoc.Canvas.Brush.Color := clHighLight;
      Rect.Bottom := Rect.Bottom + 1;
      sgrXMLDoc.Canvas.FillRect(Rect);
      sgrXMLDoc.Canvas.TextOut(Rect.Left + 2, Y, sgrXMLDoc.Cells[ACol, ARow]);
      Rect.Bottom := Rect.Bottom - 1;
    end;
  end;
  if gdFocused in State
    then DrawFocusRect(sgrXMLDoc.Canvas.Handle, Rect);
  sgrXMLDoc.DefaultDrawing := true;
end;

procedure TQImport3WizardF.sgrXMLDocMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: integer;
begin
  sgrXMLDoc.MouseToCell(X, Y, ACol, ARow);
  if not Assigned(lvXMLDocFields.Selected) then Exit;
  if XMLDocCol = ACol + 1
    then lvXMLDocFields.Selected.SubItems[0] := EmptyStr
    else lvXMLDocFields.Selected.SubItems[0] := IntToStr(ACol + 1);
  lvXMLDocFieldsChange(lvXMLDocFields, lvXMLDocFields.Selected, ctState);
end;

procedure TQImport3WizardF.lvXMLDocFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;
  cbXMLDocColNumber.OnChange := nil;
  try
    if Item.SubItems.Count > 0 then
      cbXMLDocColNumber.ItemIndex := cbXMLDocColNumber.Items.IndexOf(Item.SubItems[0]);
    sgrXMLDoc.Repaint;
    TuneButtons;
  finally
    cbXMLDocColNumber.OnChange := cbXMLDocColNumberChange;
  end;
end;

procedure TQImport3WizardF.tbtXMLDocAutoFillClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvXMLDocFields.Items.Count - 1 do
    if (i <= sgrXMLDoc.ColCount - 2) then
      lvXMLDocFields.Items[i].SubItems[0] := IntToStr(i + 2)
    else
      lvXMLDocFields.Items[i].SubItems[0] := EmptyStr;
  lvXMLDocFieldsChange(lvXMLDocFields, lvXMLDocFields.Selected, ctState);
end;

procedure TQImport3WizardF.tbtXMLDocClearClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvXMLDocFields.Items.Count - 1 do
    lvXMLDocFields.Items[i].SubItems[0] := EmptyStr;
  lvXMLDocFieldsChange(lvXMLDocFields, lvXMLDocFields.Selected, ctState);
end;

procedure TQImport3WizardF.bXMLDocFillGridClick(Sender: TObject);
begin
  XMLDocDataLocation := TXMLDataLocation(cbXMLDocDataLocation.ItemIndex);
  XMLDocXPath := edXMLDocXPath.Text;
end;

procedure TQImport3WizardF.bXMLDocBuildTreeClick(Sender: TObject);
var
  F: TForm;
  Start, Finish: TDateTime;
begin
  if not FNeedLoadFile then
  begin
    Start := Now;
    F := ShowLoading(Self, 'Build TreeView');
    try
      Application.ProcessMessages;
      XMLFile2TreeView(tvXMLDoc, FXMLDocFile);
     finally
      Finish := Now;
      while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
        Finish := Now;
      F.Free;
    end;
  end;
end;

function TQImport3WizardF.GetXPath: qiString;
var
  CurrentNode: TTreeNode;
begin
  Result := '/';
  if Assigned(tvXMLDoc.Selected) then
  begin
    CurrentNode := tvXMLDoc.Selected;
    if Assigned(CurrentNode) then
      if not Assigned(CurrentNode.Data) then
        Result := Concat('/', CurrentNode.Text)
      else
      if Integer(CurrentNode.Data) = 2 then
        Result := Concat('[', CurrentNode.Text, ']');

    while CurrentNode.Parent <> nil do
    begin
      CurrentNode := CurrentNode.Parent;
      Result := Concat('/', CurrentNode.Text, Result);
    end;
  end;
end;

procedure TQImport3WizardF.bXMLDocGetXPathClick(Sender: TObject);
begin
  edXMLDocXPath.Text := GetXPath;
end;
{$ENDIF}
{$ENDIF}

{$IFDEF XLSX}
{$IFDEF VCL6}
// Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx
// Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx
// Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx Xlsx

procedure TQImport3WizardF.edXlsxSkipRowsChange(Sender: TObject);
begin
  XlsxSkipLines := StrToIntDef(edXlsxSkipRows.Text, 0);
end;

procedure TQImport3WizardF.lvXlsxFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;
  if Assigned(FXlsxCurrentStringGrid) then
  begin
    FXlsxCurrentStringGrid.Repaint;
    TuneButtons;
  end;
end;

procedure TQImport3WizardF.sgrXlsxDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  X, Y: Integer;
  grid: TqiStringGrid;
begin
  if not (Sender is TqiStringGrid) then Exit;
  grid := Sender as TqiStringGrid;

  X := Rect.Left + (Rect.Right - Rect.Left -
   grid.Canvas.TextWidth(grid.Cells[ACol, ARow])) div 2;
  Y := Rect.Top + (Rect.Bottom - Rect.Top -
    grid.Canvas.TextHeight(grid.Cells[ACol, ARow])) div 2;
  if gdFixed in State then
  begin
    if (ACol = XlsxCol) and (ARow = 0) then
      grid.Canvas.Font.Style := grid.Canvas.Font.Style + [fsBold]
    else
      grid.Canvas.Font.Style := grid.Canvas.Font.Style - [fsBold];
    grid.Canvas.FillRect(Rect);
    grid.Canvas.TextOut(X - 1, Y + 1, grid.Cells[ACol, ARow]);
  end
  else begin
    grid.DefaultDrawing := false;
    grid.Canvas.Brush.Color := clWindow;
    grid.Canvas.FillRect(Rect);
    grid.Canvas.Font.Color := clWindowText;
    grid.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2,
      grid.Cells[ACol, ARow]);

    if (ACol = XlsxCol) and (ARow > 0) and
      ((XlsxSheetNamingCol = '') or (XlsxSheetNamingCol = XlsxSheetName)) then
    begin
      grid.Canvas.Font.Color := clHighLightText;
      grid.Canvas.Brush.Color := clHighLight;
      Rect.Bottom := Rect.Bottom + 1;
      grid.Canvas.FillRect(Rect);
      grid.Canvas.TextOut(Rect.Left + 2, Y, grid.Cells[ACol, ARow]);
      Rect.Bottom := Rect.Bottom - 1;
    end;
  end;
  if gdFocused in State
    then DrawFocusRect(grid.Canvas.Handle, Rect);
  grid.DefaultDrawing := true;
end;

procedure TQImport3WizardF.sgrXlsxMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

  var
    ACol, ARow: Integer;
    grid: TqiStringGrid;
begin
  if not (Sender is TqiStringGrid) then Exit;
  grid := Sender as TqiStringGrid;

  grid.MouseToCell(X, Y, ACol, ARow);
  if not Assigned(lvXlsxFields.Selected) then Exit;
  if (ACol = XlsxCol) and (XlsxSheetNamingCol = XlsxSheetName) then
    lvXlsxFields.Selected.SubItems[0] := EmptyStr
  else
    lvXlsxFields.Selected.SubItems[0] := SheetPrefix + XlsxSheetName +
      SheetPostfix + GetColNameFromIndex(ACol) + IntToStr(BeginRowInt) +
      LeftRightRangesDelim + ColFinishFlag;

  lvXlsxFieldsChange(lvXlsxFields, lvXlsxFields.Selected, ctState);
end;

procedure TQImport3WizardF.tlbXlsxAutoFillClick(Sender: TObject);
var
  i: Integer;
begin
  if Assigned(FXlsxCurrentStringGrid) then
  begin
    for i := 0 to lvXlsxFields.Items.Count - 1 do
      if (i <= FXlsxCurrentStringGrid.ColCount - 2) then
        lvXlsxFields.Items[i].SubItems[0] := SheetPrefix + SheetIndexPrefix +
          IntToStr(pcXlsxFile.ActivePageIndex + 1) + SheetPostfix +
          GetColNameFromIndex(i + 1) + IntToStr(BeginRowInt) +
          LeftRightRangesDelim + ColFinishFlag
      else
        lvXlsxFields.Items[i].SubItems[0] := EmptyStr;
    lvXlsxFieldsChange(lvXlsxFields, lvXlsxFields.Selected, ctState);
  end;
end;

procedure TQImport3WizardF.tlbXlsxClearClick(Sender: TObject);
begin
  XlsxClearList;
end;

procedure TQImport3WizardF.pcXlsxFileChange(Sender: TObject);
begin
  XlsxSheetName := pcXlsxFile.ActivePage.Caption;
end;

procedure TQImport3WizardF.XlsxTune;
begin
  ShowTip(tsXlsxOptions, 6, 21, 39, tsXlsxOptions.Width - 6, QImportLoadStr(QIW_Xlsx_Tip));

  pgImport.ActivePage := tsXlsxOptions;
  laStep_11.Caption := Format(QImportLoadStr(QIW_Step), [1, 3]);

  if FNeedLoadFile then XlsxFillGrid;

  if (lvXlsxFields.Items.Count > 0) and
     not Assigned(lvXlsxFields.Selected) then
  begin
    lvXlsxFields.Items[0].Focused := True;
    lvXlsxFields.Items[0].Selected := True;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.XlsxFillList;
var
  i: Integer;
begin
  if not (aiXlsx in Wizard.AllowedImports) then Exit;
  XlsxClearList;
  for i := 0 to QImportDestinationColCount(False, ImportDestination, DataSet,
                  DBGrid, ListView, StringGrid) - 1 do
    with lvXlsxFields.Items.Add do
    begin
      Caption := QImportDestinationColName(False, ImportDestination, DataSet,
        DBGrid, Self.ListView, StringGrid, GridCaptionRow, i);
      SubItems.Add(EmptyStr);
    end;

  if lvXlsxFields.Items.Count > 0 then
  begin
    lvXlsxFields.Items[0].Focused := True;
    lvXlsxFields.Items[0].Selected := True;
  end;
end;

procedure TQImport3WizardF.XlsxClearList;
var
  i: Integer;
begin
  for i := 0 to lvXlsxFields.Items.Count - 1 do
    lvXlsxFields.Items[i].SubItems[0] := EmptyStr;
  lvXlsxFieldsChange(lvXlsxFields, lvXlsxFields.Selected, ctState);
end;

procedure TQImport3WizardF.XlsxFillGrid;
var
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  XlsxFile: TXlsxFile;
  Start, Finish: TDateTime;
  F: TForm;
  i, j: Integer;
begin
  if not FileExists(FileName) then Exit;

  XlsxFile := TXlsxFile.Create;
  try
    XlsxFile.FileName := FFileName;
    XlsxFile.Workbook.NeedFillMerge := XlsxNeedFillMerge;
    XlsxFile.Workbook.LoadHiddenSheets := XlsxLoadHiddenSheet;
    Start := Now;
    F := ShowLoading(Self, FFileName);
    try
      Application.ProcessMessages;
      XlsxFile.Load;
      XlsxClearAll;
      for i := 0 to XlsxFile.Workbook.WorkSheets.Count - 1 do
      begin
        TabSheet := TTabSheet.Create(pcXlsxFile);
        TabSheet.PageControl := pcXlsxFile;
        TabSheet.Caption := XlsxFile.Workbook.WorkSheets[i].Name;

        StringGrid := TqiStringGrid.Create(TabSheet);
        StringGrid.Parent := TabSheet;
        StringGrid.Align := alClient;
        StringGrid.ColCount := 257;
        StringGrid.RowCount := 257;
        StringGrid.FixedCols := 1;
        StringGrid.FixedRows := 1;
        StringGrid.DefaultColWidth := 64;
        StringGrid.DefaultRowHeight := 16;
        StringGrid.ColWidths[0] := 30;
        StringGrid.Tag := i;
        StringGrid.Options := StringGrid.Options - [goRangeSelect];
        GridFillFixedCells(StringGrid);
        StringGrid.OnDrawCell := sgrXlsxDrawCell;
        StringGrid.OnMouseDown := sgrXlsxMouseDown;
        for j := 0 to XlsxFile.Workbook.WorkSheets[i].Cells.Count - 1 do
          StringGrid.Cells[XlsxFile.Workbook.WorkSheets[i].Cells[j].Col,
            XlsxFile.Workbook.WorkSheets[i].Cells[j].Row] := XlsxFile.Workbook.WorkSheets[i].Cells[j].Value;
        if (i = 0) then
          XlsxSheetName := TabSheet.Caption;
      end;
    finally
      Finish := Now;
      while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
        Finish := Now;
      if Assigned(F) then
        F.Free;
    end;
  finally
    XlsxFile.Free;
  end;
  FNeedLoadFile := false;  
end;

procedure TQImport3WizardF.XlsxClearAll;
begin
  FXlsxCurrentStringGrid := nil;
  while pcXlsxFile.PageCount > 0 do
    pcXlsxFile.Pages[0].Free;
  XlsxSheetName := '';
end;

function TQImport3WizardF.XlsxReady: boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to lvXlsxFields.Items.Count - 1 do
  begin
    Result := lvXlsxFields.Items[i].SubItems[0] <> '';
    if Result then Break;
  end;
end;

function TQImport3WizardF.XlsxCol: Integer;
  var
    TmpValue: string;
  XlsxRange: TXlsxMapParserItem;
begin
  Result := 0;
  if Assigned(lvXlsxFields.Selected) then
  begin
    TmpValue := lvXlsxFields.Selected.SubItems[0];
    if TmpValue <> EmptyStr then
    begin
      XlsxRange := TXlsxMapParserItem.Create(TmpValue, XlsxSkipLines, 0);
      try
        Result := XlsxRange.Ranges[0].FirstCol;
      finally
        XlsxRange.Free;
      end;
    end;
  end;
end;

function TQImport3WizardF.XlsxSheetNamingCol: qiString;
  var
    TmpValue: string;
    XlsxRange: TXlsxMapParserItem;
    TmpSheetIndex: Integer;
begin
  Result := '';
  if Assigned(lvXlsxFields.Selected) then
  begin
    TmpValue := lvXlsxFields.Selected.SubItems[0];
    if TmpValue <> EmptyStr then
    begin
      XlsxRange := TXlsxMapParserItem.Create(TmpValue, XlsxSkipLines, 0);
      try
        Result := XlsxRange.Ranges[0].SheetName;
        if Result = '' then
        begin
          TmpSheetIndex := XlsxRange.Ranges[0].SheetIndex;

          if (TmpSheetIndex <= pcXlsxFile.PageCount) and (TmpSheetIndex > 0) then
            Result := pcXlsxFile.Pages[TmpSheetIndex - 1].Caption;
        end;
      finally
        XlsxRange.Free;
      end;
    end;
  end;
end;

procedure TQImport3WizardF.SetXlsxSkipLines(const Value: Integer);
begin
  if FXlsxSkipLines <> Value then
  begin
    FXlsxSkipLines := Value;
    edXlsxSkipRows.Text := IntToStr(FXlsxSkipLines);
  end;
end;

procedure TQImport3WizardF.SetXlsxSheetName(const Value: qiString);
var
  i: Integer;
begin
  for i := 0 to pcXlsxFile.PageCount - 1 do
    if pcXlsxFile.Pages[i].Caption = Value then
    begin
      pcXlsxFile.ActivePage := pcXlsxFile.Pages[i];
      Break;
    end;
  if Assigned(pcXlsxFile.ActivePage) then
  begin
    FXlsxSheetName := pcXlsxFile.ActivePage.Caption;
    if pcXlsxFile.ActivePage.ComponentCount > 0 then
      if pcXlsxFile.ActivePage.Components[0] is TqiStringGrid then
        FXlsxCurrentStringGrid := TqiStringGrid(pcXlsxFile.ActivePage.Components[0]);
  end;
end;

procedure TQImport3WizardF.SetXlsxLoadHiddenSheet(const Value: Boolean);
begin
  FXlsxLoadHiddenSheet := Value;
end;

procedure TQImport3WizardF.SetXlsxNeedFillMerge(const Value: Boolean);
begin
  FXlsxNeedFillMerge := Value;
end;

{$ENDIF}
{$ENDIF}


{$IFDEF DOCX}
{$IFDEF VCL6}
// Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx
// Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx
// Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx Docx

procedure TQImport3WizardF.edDocxSkipRowsChange(Sender: TObject);
begin
  DocxSkipLines := StrToIntDef(edDocxSkipRows.Text, 0);
end;

procedure TQImport3WizardF.lvDocxFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;
  if Assigned(FDocxCurrentStringGrid) then
  begin
    FDocxCurrentStringGrid.Repaint;
    TuneButtons;
  end;
end;

procedure TQImport3WizardF.sgrDocxDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  X, Y: Integer;
  grid: TqiStringGrid;
begin
  if not (Sender is TqiStringGrid) then Exit;
  grid := Sender as TqiStringGrid;

  X := Rect.Left + (Rect.Right - Rect.Left -
   grid.Canvas.TextWidth(grid.Cells[ACol, ARow])) div 2;
  Y := Rect.Top + (Rect.Bottom - Rect.Top -
    grid.Canvas.TextHeight(grid.Cells[ACol, ARow])) div 2;
  if gdFixed in State then
  begin
    grid.Canvas.FillRect(Rect);
    grid.Canvas.TextOut(X - 1, Y + 1, grid.Cells[ACol, ARow]);
  end
  else begin
    grid.DefaultDrawing := False;
    grid.Canvas.Brush.Color := clWindow;
    grid.Canvas.FillRect(Rect);
    grid.Canvas.Font.Color := clWindowText;

    if Boolean(grid.Objects[ACol, ARow]) = True then
      DrawBlob(TqiStringGrid(Sender).Canvas, Rect)
    else
      grid.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2, grid.Cells[ACol, ARow]);

    if (ACol = DocxCol - 1) then
    begin
      grid.Canvas.Font.Color := clHighLightText;
      grid.Canvas.Brush.Color := clHighLight;
      Rect.Bottom := Rect.Bottom + 1;
      grid.Canvas.FillRect(Rect);
      if Boolean(grid.Objects[ACol, ARow]) = True then
        DrawBlob(TqiStringGrid(Sender).Canvas, Rect)
      else
        grid.Canvas.TextOut(Rect.Left + 2, Y, grid.Cells[ACol, ARow]);
      Rect.Bottom := Rect.Bottom - 1;
    end;
  end;
  if gdFocused in State then
    DrawFocusRect(grid.Canvas.Handle, Rect);
  grid.DefaultDrawing := True;

end;

procedure TQImport3WizardF.sgrDocxMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: Integer;
  grid: TqiStringGrid;
begin
  if not (Sender is TqiStringGrid) then Exit;
  grid := Sender as TqiStringGrid;

  grid.MouseToCell(X, Y, ACol, ARow);
  if not Assigned(lvDocxFields.Selected) then Exit;
  if DocxCol = ACol + 1 then
    lvDocxFields.Selected.SubItems[0] := EmptyStr
  else
    if ACol > -1 then
      lvDocxFields.Selected.SubItems[0] := IntToStr(ACol + 1);

  lvDocxFieldsChange(lvDocxFields, lvDocxFields.Selected, ctState);
end;

procedure TQImport3WizardF.tlbDocxAutoFillClick(Sender: TObject);
var
  i: Integer;
begin
  if Assigned(FDocxCurrentStringGrid) then
  begin
    for i := 0 to lvDocxFields.Items.Count - 1 do
      if (i <= FDocxCurrentStringGrid.ColCount - 1) then
        lvDocxFields.Items[i].SubItems[0] := IntToStr(i + 1)
      else
        lvDocxFields.Items[i].SubItems[0] := EmptyStr;
    lvDocxFieldsChange(lvDocxFields, lvDocxFields.Selected, ctState);
  end;
end;

procedure TQImport3WizardF.tlbDocxClearClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvDocxFields.Items.Count - 1 do
    lvDocxFields.Items[i].SubItems[0] := EmptyStr;
  lvDocxFieldsChange(lvDocxFields, lvDocxFields.Selected, ctState);
end;

procedure TQImport3WizardF.pcDocxFileChange(Sender: TObject);
begin
  DocxTableNumber := pcDocxFile.ActivePage.Tag;
end;

procedure TQImport3WizardF.DocxTune;
begin
  ShowTip(tsDocxOptions, 6, 21, 39, tsDocxOptions.Width - 6, QImportLoadStr(QIW_Docx_Tip));
  
  pgImport.ActivePage := tsDocxOptions;
  laStep_12.Caption := Format(QImportLoadStr(QIW_Step), [1, 3]);

  if FNeedLoadFile then DocxFillGrid;

  if (lvDocxFields.Items.Count > 0) and
    not Assigned(lvDocxFields.Selected) then
  begin
    lvDocxFields.Items[0].Focused := True;
    lvDocxFields.Items[0].Selected := True;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.DocxFillList;
var
  i: Integer;
begin
  if not (aiDocx in Wizard.AllowedImports) then Exit;
  DocxClearList;
  for i := 0 to QImportDestinationColCount(False, ImportDestination, DataSet,
                  DBGrid, ListView, StringGrid) - 1 do
    with lvDocxFields.Items.Add do
    begin
      Caption := QImportDestinationColName(False, ImportDestination, DataSet,
        DBGrid, Self.ListView, StringGrid, GridCaptionRow, i);
      SubItems.Add(EmptyStr);
    end;
    
  if lvDocxFields.Items.Count > 0 then
  begin
    lvDocxFields.Items[0].Focused := True;
    lvDocxFields.Items[0].Selected := True;
  end;
end;

procedure TQImport3WizardF.DocxClearList;
var
  i: Integer;
begin
  for i := lvDocxFields.Items.Count - 1 downto 0 do
    lvDocxFields.Items.Delete(i);
end;

procedure TQImport3WizardF.DocxFillGrid;
var
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  DocxFile: TDocxFile;
  F: TForm;
  Start, Finish: TDateTime;
  i, j, k:  Integer;
begin
  if not FileExists(FileName) then Exit;

  DocxFile := TDocxFile.Create;
  try
    DocxFile.FileName := FFileName;
//    DocxFile.NeedFillMerge := Import.NeedFillMerge;
    Start := Now;
    F := ShowLoading(Self, FFileName);
    try
      Application.ProcessMessages;
      DocxFile.Load;
      DocxClearAll;
      for i := 0 to DocxFile.Tables.Count - 1 do
      begin
        TabSheet := TTabSheet.Create(pcDocxFile);
        TabSheet.PageControl := pcDocxFile;
        TabSheet.Caption := 'Table ' + IntToStr(Succ(i));
        TabSheet.Tag := Succ(i);

        StringGrid := TqiStringGrid.Create(TabSheet);
        StringGrid.Parent := TabSheet;
        StringGrid.Align := alClient;
        StringGrid.DefaultRowHeight := 16;
        StringGrid.ColCount := 1;
        StringGrid.RowCount := 1;
        StringGrid.FixedCols := 0;
        StringGrid.FixedRows := 0;
        StringGrid.Options := StringGrid.Options - [goRangeSelect];
        StringGrid.OnDrawCell := sgrDocxDrawCell;
        StringGrid.OnMouseDown := sgrDocxMouseDown;

        StringGrid.ColCount := 1;
        StringGrid.RowCount := Min(DocxFile.Tables[i].Rows.Count, 30);
        for j := 0 to StringGrid.RowCount - 1 do
        begin
          if StringGrid.ColCount < DocxFile.Tables[i].Rows[j].Cols.Count then
            StringGrid.ColCount := DocxFile.Tables[i].Rows[j].Cols.Count;
          for k := 0 to DocxFile.Tables[i].Rows[j].Cols.Count - 1 do
          begin
            StringGrid.Objects[k,j] := TObject(DocxFile.Tables[i].Rows[j].Cols[k].DisplayAsIcon);
            StringGrid.Cells[k, j] := DocxFile.Tables[i].Rows[j].Cols[k].Text;
          end;
        end;
        if (i = 0) then
          DocxTableNumber := TabSheet.Tag;
      end;
    finally
      Finish := Now;
      while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
        Finish := Now;
      if Assigned(F) then
        F.Free;
    end;
  finally
    DocxFile.Free;
  end;
  FNeedLoadFile := false;
end;

procedure TQImport3WizardF.DocxClearAll;
begin
  FDocxCurrentStringGrid := nil;
  while pcDocxFile.PageCount > 0 do
    pcDocxFile.Pages[0].Free;
  DocxTableNumber := 0;
end;

function TQImport3WizardF.DocxReady: boolean;
var
  i: Integer;
begin
  Result := False;
  if (pcDocxFile.PageCount > 0) then
    for i := 0 to lvDocxFields.Items.Count - 1 do
    begin
      Result := StrToIntDef(lvDocxFields.Items[i].SubItems[0], 0) > 0;
      if Result then Break;
    end;
end;

function TQImport3WizardF.DocxCol: Integer;
begin
  Result := 0;
  if Assigned(lvDocxFields.Selected) then
    Result := StrToIntDef(lvDocxFields.Selected.SubItems[0], 0);
end;

procedure TQImport3WizardF.SetDocxSkipLines(const Value: Integer);
begin
  if FDocxSkipLines <> Value then
  begin
    FDocxSkipLines := Value;
    edDocxSkipRows.Text := IntToStr(FDocxSkipLines);
  end;
end;

procedure TQImport3WizardF.SetDocxTableNumber(const Value: Integer);
var
  i: Integer;
begin
  for i := 0 to pcDocxFile.PageCount - 1 do
    if Value = pcDocxFile.Pages[i].Tag then
    begin
      pcDocxFile.ActivePage := pcDocxFile.Pages[i];
      Break;
    end;
  if Assigned(pcDocxFile.ActivePage) then
  begin
    FDocxTableNumber := pcDocxFile.ActivePage.Tag;
    if pcDocxFile.ActivePage.ComponentCount > 0 then
      if pcDocxFile.ActivePage.Components[0] is TqiStringGrid then
        FDocxCurrentStringGrid := TqiStringGrid(pcDocxFile.ActivePage.Components[0]);
  end;
end;

{$ENDIF}
{$ENDIF}


{$IFDEF ODS}
{$IFDEF VCL6}
// ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS
// ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS
// ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS ODS

procedure TQImport3WizardF.edODSSkipRowsChange(Sender: TObject);
begin
  ODSSkipLines := StrToIntDef(edODSSkipRows.Text, 0);
end;

procedure TQImport3WizardF.lvODSFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;
  if Assigned(FODSCurrentStringGrid) then
  begin
    FODSCurrentStringGrid.Repaint;
    TuneButtons;
  end;
end;

procedure TQImport3WizardF.sgrODSDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  X, Y: Integer;
  grid: TqiStringGrid;
begin
  if not (Sender is TqiStringGrid) then Exit;
  grid := Sender as TqiStringGrid;

  X := Rect.Left + (Rect.Right - Rect.Left -
    grid.Canvas.TextWidth(grid.Cells[ACol, ARow])) div 2;
  Y := Rect.Top + (Rect.Bottom - Rect.Top -
    grid.Canvas.TextHeight(grid.Cells[ACol, ARow])) div 2;
  if gdFixed in State then
  begin
    if (ACol = ODSCol) and (ARow = 0) then
      grid.Canvas.Font.Style := grid.Canvas.Font.Style + [fsBold]
    else
      grid.Canvas.Font.Style := grid.Canvas.Font.Style - [fsBold];
    grid.Canvas.FillRect(Rect);
    grid.Canvas.TextOut(X - 1, Y + 1, grid.Cells[ACol, ARow]);
  end
  else begin
    grid.DefaultDrawing := false;
    grid.Canvas.Brush.Color := clWindow;
    grid.Canvas.FillRect(Rect);
    grid.Canvas.Font.Color := clWindowText;
    grid.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2,
      grid.Cells[ACol, ARow]);

    if (ACol = ODSCol) and (ARow > 0) then
    begin
      grid.Canvas.Font.Color := clHighLightText;
      grid.Canvas.Brush.Color := clHighLight;
      Rect.Bottom := Rect.Bottom + 1;
      grid.Canvas.FillRect(Rect);
      grid.Canvas.TextOut(Rect.Left + 2, Y, grid.Cells[ACol, ARow]);
      Rect.Bottom := Rect.Bottom - 1;
    end;
  end;
  if gdFocused in State
    then DrawFocusRect(grid.Canvas.Handle, Rect);
  grid.DefaultDrawing := true;
end;

procedure TQImport3WizardF.sgrODSMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: Integer;
  grid: TqiStringGrid;
begin
  if not (Sender is TqiStringGrid) then Exit;
  grid := Sender as TqiStringGrid;

  grid.MouseToCell(X, Y, ACol, ARow);
  if not Assigned(lvODSFields.Selected) then Exit;
  if ACol = 0 then Exit;
  if ODSCol = ACol
    then lvODSFields.Selected.SubItems[0] := EmptyStr
    else lvODSFields.Selected.SubItems[0] := Col2Letter(ACol);
  lvODSFieldsChange(lvODSFields, lvODSFields.Selected, ctState);
end;

procedure TQImport3WizardF.tlbODSAutoFillClick(Sender: TObject);
var
  i: Integer;
begin
  if Assigned(FODSCurrentStringGrid) then
  begin
    for i := 0 to lvODSFields.Items.Count - 1 do
      if (i <= FODSCurrentStringGrid.ColCount - 2) then
        lvODSFields.Items[i].SubItems[0] := Col2Letter(i + 1)
      else
        lvODSFields.Items[i].SubItems[0] := EmptyStr;
    lvODSFieldsChange(lvODSFields, lvODSFields.Selected, ctState);
  end;
end;

procedure TQImport3WizardF.tlbODSClearClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvODSFields.Items.Count - 1 do
    lvODSFields.Items[i].SubItems[0] := EmptyStr;
  lvODSFieldsChange(lvODSFields, lvODSFields.Selected, ctState);
end;

procedure TQImport3WizardF.pcODSFileChange(Sender: TObject);
begin
  ODSSheetName := AnsiString(pcODSFile.ActivePage.Caption);
end;

procedure TQImport3WizardF.ODSTune;
begin
  ShowTip(tsODSOptions, 6, 21, 39, tsODSOptions.Width - 6, QImportLoadStr(QIW_ODS_Tip));

  pgImport.ActivePage := tsODSOptions;
  laStep_13.Caption := Format(QImportLoadStr(QIW_Step), [1, 3]);

  if FNeedLoadFile then ODSFillGrid;

  if (lvODSFields.Items.Count > 0) and not Assigned(lvODSFields.Selected) then
  begin
    lvODSFields.Items[0].Focused := True;
    lvODSFields.Items[0].Selected := True;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.ODSFillList;
var
  i: Integer;
begin
  if not (aiODS in Wizard.AllowedImports) then Exit;
  ODSClearList;
  for i := 0 to QImportDestinationColCount(False, ImportDestination, DataSet,
                  DBGrid, ListView, StringGrid) - 1 do
    with lvODSFields.Items.Add do
    begin
      Caption := QImportDestinationColName(False, ImportDestination, DataSet,
        DBGrid, Self.ListView, StringGrid, GridCaptionRow, i);
      SubItems.Add(EmptyStr);
    end;

  if lvODSFields.Items.Count > 0 then
  begin
    lvODSFields.Items[0].Focused := True;
    lvODSFields.Items[0].Selected := True;
  end;
end;

procedure TQImport3WizardF.ODSClearList;
var
  i: Integer;
begin
  for i := lvODSFields.Items.Count - 1 downto 0 do
    lvODSFields.Items.Delete(i);
end;

procedure TQImport3WizardF.ODSFillGrid;
var
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  k, i, j: integer;
  F: TForm;
  Start, Finish: TDateTime;
  W: integer;
  FODSFile: TODSFile;
begin
  if not FileExists(FFileName) then Exit;
  FODSFile := TODSFile.Create;
  try
    FODSFile.FileName := FFileName;
    Start := Now;
    F := ShowLoading(Self, FFileName);
    try
      Application.ProcessMessages;
      FODSFile.Workbook.IsNotExpanding := true;
      FODSFile.Load;
      ODSClearAll;
      for k := 0 to FODSFile.Workbook.SpreadSheets.Count - 1 do
      begin
        TabSheet := TTabSheet.Create(pcODSFile);
        TabSheet.PageControl := pcODSFile;
        TabSheet.Caption := string(FODSFile.Workbook.SpreadSheets[k].Name);
        TabSheet.Tag := k;

        StringGrid := TqiStringGrid.Create(TabSheet);
        StringGrid.Parent := TabSheet;
        StringGrid.Align := alClient;
        StringGrid.ColCount := 257;
        StringGrid.RowCount := 257;
        StringGrid.FixedCols := 1;
        StringGrid.FixedRows := 1;
        StringGrid.Tag := k;
        StringGrid.DefaultColWidth := 64;
        StringGrid.DefaultRowHeight := 16;
        StringGrid.ColWidths[0] := 30;
        StringGrid.Options := StringGrid.Options - [goRangeSelect];
        GridFillFixedCells(StringGrid);
        StringGrid.OnDrawCell := sgrODSDrawCell;
        StringGrid.OnMouseDown := sgrODSMouseDown;

        for i := 0 to FODSFile.Workbook.SpreadSheets[k].DataCells.RowCount - 1 do
          for j := 0 to FODSFile.Workbook.SpreadSheets[k].DataCells.ColCount - 1 do
          begin
            StringGrid.Cells[j + 1, i + 1] :=
              FODSFile.Workbook.SpreadSheets[k].DataCells.Cells[j, i];
            W := StringGrid.Canvas.TextWidth(StringGrid.Cells[j + 1, i + 1]);
            if W + 10 > StringGrid.ColWidths[j + 1] then
              if W + 10 < 130 then
                StringGrid.ColWidths[j + 1] := W + 10
              else
                StringGrid.ColWidths[j + 1] := 130;
          end;
         if (k = 0) then
           ODSSheetName := AnsiString(TabSheet.Caption);
      end;
    finally
      Finish := Now;
      while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
        Finish := Now;
      if Assigned(F) then
        F.Free;
    end;
  finally
    if assigned(FODSFile) then
      FODSFile.Free;
  end;
  FNeedLoadFile := false;
end;

procedure TQImport3WizardF.ODSClearAll;
begin
  FODSCurrentStringGrid := nil;
  while pcODSFile.PageCount > 0 do
    pcODSFile.Pages[0].Free;
  ODSSheetName := '';
end;

function TQImport3WizardF.ODSReady: boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to lvODSFields.Items.Count - 1 do
  begin
    Result := lvODSFields.Items[i].SubItems[0] <> '';
    if Result then Break;
  end;
end;

function TQImport3WizardF.ODSCol: Integer;
begin
  Result := 0;
  if Assigned(lvODSFields.Selected) then
    Result := GetColIdFromColIndex(lvODSFields.Selected.SubItems[0]);
end;

procedure TQImport3WizardF.SetODSSkipLines(const Value: Integer);
begin
  if FODSSkipLines <> Value then
  begin
    FODSSkipLines := Value;
    edODSSkipRows.Text := IntToStr(FODSSkipLines);
  end;
end;

procedure TQImport3WizardF.SetODSSheetName(const Value: AnsiString);
var
  i: Integer;
begin
  for i := 0 to pcODSFile.PageCount - 1 do
    if pcODSFile.Pages[i].Caption = string(Value) then
    begin
      pcODSFile.ActivePage := pcODSFile.Pages[i];
      Break;
    end;
  if Assigned(pcODSFile.ActivePage) then
  begin
    FODSSheetName := AnsiString(pcODSFile.ActivePage.Caption);
    if pcODSFile.ActivePage.ComponentCount > 0 then
      if pcODSFile.ActivePage.Components[0] is TqiStringGrid then
        FODSCurrentStringGrid := TqiStringGrid(pcODSFile.ActivePage.Components[0]);
  end;
end;

{$ENDIF}
{$ENDIF}

{$IFDEF ODT}
{$IFDEF VCL6}
// ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT
// ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT
// ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT ODT

procedure TQImport3WizardF.edODTSkipRowsChange(Sender: TObject);
begin
  ODTSkipLines := StrToIntDef(edODTSkipRows.Text, 0);
end;

procedure TQImport3WizardF.lvODTFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;
  if Assigned(FODTCurrentStringGrid) then
  begin
    FODTCurrentStringGrid.Repaint;
    TuneButtons;
  end;
end;

procedure TQImport3WizardF.sgrODTDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  X, Y: Integer;
  grid: TqiStringGrid;
begin
  if not (Sender is TqiStringGrid) then Exit;
  grid := Sender as TqiStringGrid;

  X := Rect.Left + (Rect.Right - Rect.Left -
    grid.Canvas.TextWidth(grid.Cells[ACol, ARow])) div 2;
  Y := Rect.Top + (Rect.Bottom - Rect.Top -
    grid.Canvas.TextHeight(grid.Cells[ACol, ARow])) div 2;
  if gdFixed in State then
  begin
    if (ACol = ODTCol) and (ARow = 0) then
      grid.Canvas.Font.Style := grid.Canvas.Font.Style + [fsBold]
    else
      grid.Canvas.Font.Style := grid.Canvas.Font.Style - [fsBold];
    grid.Canvas.FillRect(Rect);
    grid.Canvas.TextOut(X - 1, Y + 1, grid.Cells[ACol, ARow]);
  end
  else begin
    grid.DefaultDrawing := false;
    grid.Canvas.Brush.Color := clWindow;
    grid.Canvas.FillRect(Rect);
    grid.Canvas.Font.Color := clWindowText;
    grid.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2,
      grid.Cells[ACol, ARow]);

    if (ACol = ODTCol) and (ARow > 0) then
    begin
      grid.Canvas.Font.Color := clHighLightText;
      grid.Canvas.Brush.Color := clHighLight;
      Rect.Bottom := Rect.Bottom + 1;
      grid.Canvas.FillRect(Rect);
      grid.Canvas.TextOut(Rect.Left + 2, Y, grid.Cells[ACol, ARow]);
      Rect.Bottom := Rect.Bottom - 1;
    end;
  end;
  if gdFocused in State
    then DrawFocusRect(grid.Canvas.Handle, Rect);
  grid.DefaultDrawing := true;
end;

procedure TQImport3WizardF.sgrODTMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: Integer;
  grid: TqiStringGrid;
begin
  if not (Sender is TqiStringGrid) then Exit;
  grid := Sender as TqiStringGrid;

  grid.MouseToCell(X, Y, ACol, ARow);
  if not Assigned(lvODTFields.Selected) then Exit;
  if ACol = 0 then Exit;
  if ODTCol = ACol
    then lvODTFields.Selected.SubItems[0] := EmptyStr
    else lvODTFields.Selected.SubItems[0] := Col2Letter(ACol);
  lvODTFieldsChange(lvODTFields, lvODTFields.Selected, ctState);
end;

procedure TQImport3WizardF.tlbODTAutoFillClick(Sender: TObject);
var
  i: Integer;
begin
  if Assigned(FODTCurrentStringGrid) then
  begin
    for i := 0 to lvODTFields.Items.Count - 1 do
      if (i <= FODTCurrentStringGrid.ColCount - 2) then
        lvODTFields.Items[i].SubItems[0] := Col2Letter(i + 1)
      else
        lvODTFields.Items[i].SubItems[0] := EmptyStr;
    lvODTFieldsChange(lvODTFields, lvODTFields.Selected, ctState);
  end;
end;

procedure TQImport3WizardF.tlbODTClearClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvODTFields.Items.Count - 1 do
    lvODTFields.Items[i].SubItems[0] := EmptyStr;
  lvODTFieldsChange(lvODTFields, lvODTFields.Selected, ctState);
end;

procedure TQImport3WizardF.pcODTFileChange(Sender: TObject);
begin
  ODTSheetName := AnsiString(pcODTFile.ActivePage.Caption);
end;

procedure TQImport3WizardF.ODTTune;
begin
  ShowTip(tsODTOptions, 6, 21, 39, tsODTOptions.Width - 6, QImportLoadStr(QIW_ODT_Tip));

  pgImport.ActivePage := tsODTOptions;
  laStep_14.Caption := Format(QImportLoadStr(QIW_Step), [1, 3]);

  if FNeedLoadFile then
    ODTFillGrid;

  if (lvODTFields.Items.Count > 0) and not Assigned(lvODTFields.Selected) then
  begin
    lvODTFields.Items[0].Focused := True;
    lvODTFields.Items[0].Selected := True;
  end;
  TuneButtons;
end;

procedure TQImport3WizardF.ODTFillList;
var
  i: Integer;
begin
  if not (aiODT in Wizard.AllowedImports) then Exit;
  ODTClearList;
  for i := 0 to QImportDestinationColCount(False, ImportDestination, DataSet,
                  DBGrid, ListView, StringGrid) - 1 do
    with lvODTFields.Items.Add do
    begin
      Caption := QImportDestinationColName(False, ImportDestination, DataSet,
        DBGrid, Self.ListView, StringGrid, GridCaptionRow, i);
      SubItems.Add(EmptyStr);
    end;

  if lvODTFields.Items.Count > 0 then
  begin
    lvODTFields.Items[0].Focused := True;
    lvODTFields.Items[0].Selected := True;
  end;
end;

procedure TQImport3WizardF.ODTClearList;
var
  i: Integer;
begin
  for i := lvODTFields.Items.Count - 1 downto 0 do
    lvODTFields.Items.Delete(i);
end;

procedure TQImport3WizardF.ODTFillGrid;
var
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  k, i, j, modif: integer;
  F: TForm;
  Start, Finish: TDateTime;
  W: integer;
  FODTFile: TODTFile;
begin
  if ODTUseHeader then
    modif := 0
  else
    modif := 1;
  if not FileExists(FFileName) then Exit;
  FODTFile := TODTFile.Create;
  try
    FODTFile.FileName := FFileName;
    Start := Now;
    F := ShowLoading(Self, FFileName);
    try
      Application.ProcessMessages;
      FODTFile.Load;
      ODTClearAll;
      SetLength(FODTComplexArray, FODTFile.Workbook.SpreadSheets.Count);
      for k := 0 to FODTFile.Workbook.SpreadSheets.Count - 1 do
      begin
        TabSheet := TTabSheet.Create(pcODTFile);
        TabSheet.PageControl := pcODTFile;
        TabSheet.Caption := FODTFile.Workbook.SpreadSheets[k].Name;
        TabSheet.Tag := k;

        StringGrid := TqiStringGrid.Create(TabSheet);
        FODTComplexArray[k] := FODTFile.Workbook.SpreadSheets[k].IsComplexTable;

        StringGrid.Parent := TabSheet;
        StringGrid.Align := alClient;
        StringGrid.ColCount := 257;
        StringGrid.RowCount := 257;
        StringGrid.FixedCols := 1;
        StringGrid.FixedRows := 1;
        StringGrid.Tag := k;
        StringGrid.DefaultColWidth := 64;
        StringGrid.DefaultRowHeight := 16;
        StringGrid.ColWidths[0] := 30;
        StringGrid.Options := StringGrid.Options - [goRangeSelect];
        GridFillFixedCells(StringGrid);
        StringGrid.OnDrawCell := sgrODTDrawCell;
        StringGrid.OnMouseDown := sgrODTMouseDown;

        for i := 0 to FODTFile.Workbook.SpreadSheets[k].DataCells.RowCount - 1 do
          for j := 0 to FODTFile.Workbook.SpreadSheets[k].DataCells.ColCount - 1 do
          begin
            StringGrid.Cells[j + 1, i + 1 * modif] :=
              FODTFile.Workbook.SpreadSheets[k].DataCells.Cells[j, i];
            W := StringGrid.Canvas.TextWidth(StringGrid.Cells[j + 1, i + 1 * modif]);
            if W + 10 > StringGrid.ColWidths[j + 1] then
              if W + 10 < 130 then
                StringGrid.ColWidths[j + 1] := W + 10
              else
                StringGrid.ColWidths[j + 1] := 130;
          end;
         if (k = 0) then
         begin
           ODTSheetName := AnsiString(TabSheet.Caption);
           ODTComplex := FODTFile.Workbook.SpreadSheets[k].IsComplexTable;
         end;
      end;
    finally
      Finish := Now;
      while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
        Finish := Now;
      if Assigned(F) then
        F.Free;
    end;
  finally
    if assigned(FODTFile) then
      FODTFile.Free;
  end;
  FNeedLoadFile := false;
end;

procedure TQImport3WizardF.ODTClearAll;
begin
  FODTCurrentStringGrid := nil;
  while pcODTFile.PageCount > 0 do
    pcODTFile.Pages[0].Free;
  ODTSheetName := '';
end;

function TQImport3WizardF.ODTReady: boolean;
var
  i: Integer;
begin
  Result := False;
  if (pcODTFile.PageCount > 0) then
    for i := 0 to lvODTFields.Items.Count - 1 do
    begin
      Result := (lvODTFields.Items[i].SubItems[0] <> '');
      if Result then Break;
    end;
end;

function TQImport3WizardF.ODTCol: Integer;
begin
  Result := 0;
  if Assigned(lvODTFields.Selected) then
    Result := GetColIdFromColIndex(lvODTFields.Selected.SubItems[0]);
end;

procedure TQImport3WizardF.SetODTSkipLines(const Value: Integer);
begin
  if FODTSkipLines <> Value then
  begin
    FODTSkipLines := Value;
    edODTSkipRows.Text := IntToStr(FODTSkipLines);
  end;
end;

procedure TQImport3WizardF.SetODTSheetName(const Value: AnsiString);
var
  i: Integer;
begin
  for i := 0 to pcODTFile.PageCount - 1 do
    if pcODTFile.Pages[i].Caption = string(Value) then
    begin
      pcODTFile.ActivePage := pcODTFile.Pages[i];
      ODTComplex := FODTComplexArray[i];
      Break;
    end;
  if Assigned(pcODTFile.ActivePage) then
  begin
    FODTSheetName := AnsiString(pcODTFile.ActivePage.Caption);
    if pcODTFile.ActivePage.ComponentCount > 0 then
      if pcODTFile.ActivePage.Components[0] is TqiStringGrid then
        FODTCurrentStringGrid := TqiStringGrid(pcODTFile.ActivePage.Components[0]);
  end;
end;

procedure TQImport3WizardF.cbODTUseHeaderClick(Sender: TObject);
var
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  k, i, j, modif, border: integer;
  F: TForm;
  Start, Finish: TDateTime;
  W: integer;
  FODTFile: TODTFile;
  TempName: AnsiString;
begin
  if cbODTUseHeader.Checked then
  begin
    ODTUseHeader := true;
    modif := 0;
  end
  else
  begin
    ODTUseHeader := false;
    modif := 1;
  end;
  if pcODTFile.PageCount = 0 then Exit;
  if not FileExists(FFileName) then Exit;
  FODTFile := TODTFile.Create;
  try
    FODTFile.FileName := FFileName;
    Start := Now;
    F := ShowLoading(Self, 'Applying changes');
    try
      Application.ProcessMessages;
      FODTFile.Load;
      border := pcODTFile.PageCount - 1;
      TempName := ODTSheetName;
      ODTSheetName := '';
      ODTClearAll;
      for k := 0 to border do
      begin
        TabSheet := TTabSheet.Create(pcODTFile);
        TabSheet.PageControl := pcODTFile;
        TabSheet.Caption := FODTFile.Workbook.SpreadSheets[k].Name;
        TabSheet.Tag := k;

        StringGrid := TqiStringGrid.Create(TabSheet);
        StringGrid.Parent := TabSheet;
        StringGrid.Align := alClient;
        StringGrid.ColCount := 257;
        StringGrid.RowCount := 257;
        StringGrid.FixedCols := 1;
        StringGrid.FixedRows := 1;
        StringGrid.Tag := k;
        StringGrid.DefaultColWidth := 64;
        StringGrid.DefaultRowHeight := 16;
        StringGrid.ColWidths[0] := 30;
        StringGrid.Options := StringGrid.Options - [goRangeSelect];
        GridFillFixedCells(StringGrid);
        if ODTUseHeader then
          for i := 0 to StringGrid.ColCount - 1 do
            StringGrid.Cells[i, 0] := EmptyStr;
        StringGrid.OnDrawCell := sgrODTDrawCell;
        StringGrid.OnMouseDown := sgrODTMouseDown;

        for i := 0 to FODTFile.Workbook.SpreadSheets[k].DataCells.RowCount - 1 do
          for j := 0 to FODTFile.Workbook.SpreadSheets[k].DataCells.ColCount - 1 do
          begin
          StringGrid.Cells[j + 1 , i + 1 * modif] :=
            FODTFile.Workbook.SpreadSheets[k].DataCells.Cells[j, i];
          W := StringGrid.Canvas.TextWidth(StringGrid.Cells[j + 1, i + 1 * modif]);
          if W + 10 > StringGrid.ColWidths[j + 1] then
            if W + 10 < 130 then
              StringGrid.ColWidths[j + 1] := W + 10
            else
              StringGrid.ColWidths[j + 1] := 130;
          end;
      end;
    ODTSheetName := TempName;
    finally
      Finish := Now;
      while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
        Finish := Now;
      if Assigned(F) then
        F.Free;
    end;
  finally
    if assigned(FODTFile) then
      FODTFile.Free;
  end;
end;

{$ENDIF}
{$ENDIF}

end.
