unit QImport3XlsxDom;
{$I QImport3VerCtrl.Inc}
interface
{$IFDEF XLSX}
{$IFDEF VCL6}
uses
  {$IFDEF VCL16}
    Xml.xmldom,
    Xml.XMLDoc,
    Xml.XMLIntf;
  {$ELSE}
    xmldom,
    XMLDoc,
    XMLIntf;
  {$ENDIF}

type

{ Forward Decls }

  IXMLWorksheetType = interface;
  IXMLDimensionType = interface;
  IXMLSheetViewsType = interface;
  IXMLSheetViewType = interface;
  IXMLSelectionType = interface;
  IXMLSheetFormatPrType = interface;
  IXMLColsType = interface;
  IXMLColType = interface;
  IXMLSheetDataType = interface;
  IXMLRowType = interface;
  IXMLCType = interface;
  IXMLIsType = interface;
  IXMLMergeCellsType = interface;
  IXMLMergeCellType = interface;
  IXMLPageMarginsType = interface;
  IXMLPageSetupType = interface;
  IXMLColType2 = interface;
  IXMLCType2 = interface;
  IXMLCType22 = interface;
  IXMLCType222 = interface;
  IXMLCType2222 = interface;
  IXMLMergeCellType2 = interface;

{ IXMLWorksheetType }

  IXMLWorksheetType = interface(IXMLNode)
    ['{3DB3EF99-B6D4-42C2-8053-B5DEEC20681D}']
    { Property Accessors }
    function Get_Xmlns: WideString;
    function Get_Dimension: IXMLDimensionType;
    function Get_SheetViews: IXMLSheetViewsType;
    function Get_SheetFormatPr: IXMLSheetFormatPrType;
    function Get_Cols: IXMLColsType;
    function Get_SheetData: IXMLSheetDataType;
    function Get_MergeCells: IXMLMergeCellsType;
    function Get_PageMargins: IXMLPageMarginsType;
    function Get_PageSetup: IXMLPageSetupType;
    procedure Set_Xmlns(Value: WideString);
    { Methods & Properties }
    property Xmlns: WideString read Get_Xmlns write Set_Xmlns;
    property Dimension: IXMLDimensionType read Get_Dimension;
    property SheetViews: IXMLSheetViewsType read Get_SheetViews;
    property SheetFormatPr: IXMLSheetFormatPrType read Get_SheetFormatPr;
    property Cols: IXMLColsType read Get_Cols;
    property SheetData: IXMLSheetDataType read Get_SheetData;
    property MergeCells: IXMLMergeCellsType read Get_MergeCells;
    property PageMargins: IXMLPageMarginsType read Get_PageMargins;
    property PageSetup: IXMLPageSetupType read Get_PageSetup;
  end;

{ IXMLDimensionType }

  IXMLDimensionType = interface(IXMLNode)
    ['{F82FA6B5-A561-4E71-9672-B5CC475DAE88}']
    { Property Accessors }
    function Get_Ref: WideString;
    procedure Set_Ref(Value: WideString);
    { Methods & Properties }
    property Ref: WideString read Get_Ref write Set_Ref;
  end;

{ IXMLSheetViewsType }

  IXMLSheetViewsType = interface(IXMLNode)
    ['{51F5E3F1-6E56-4278-AF8D-E5017B2D7E62}']
    { Property Accessors }
    function Get_SheetView: IXMLSheetViewType;
    { Methods & Properties }
    property SheetView: IXMLSheetViewType read Get_SheetView;
  end;

{ IXMLSheetViewType }

  IXMLSheetViewType = interface(IXMLNode)
    ['{80E7E459-A6B2-478D-AE8B-17C8386BC45A}']
    { Property Accessors }
    function Get_TabSelected: Integer;
    function Get_WorkbookViewId: Integer;
    function Get_Selection: IXMLSelectionType;
    procedure Set_TabSelected(Value: Integer);
    procedure Set_WorkbookViewId(Value: Integer);
    { Methods & Properties }
    property TabSelected: Integer read Get_TabSelected write Set_TabSelected;
    property WorkbookViewId: Integer read Get_WorkbookViewId write Set_WorkbookViewId;
    property Selection: IXMLSelectionType read Get_Selection;
  end;

{ IXMLSelectionType }

  IXMLSelectionType = interface(IXMLNode)
    ['{BDFADC08-208D-411D-AD64-4A354759CBD3}']
    { Property Accessors }
    function Get_ActiveCell: WideString;
    function Get_Sqref: WideString;
    procedure Set_ActiveCell(Value: WideString);
    procedure Set_Sqref(Value: WideString);
    { Methods & Properties }
    property ActiveCell: WideString read Get_ActiveCell write Set_ActiveCell;
    property Sqref: WideString read Get_Sqref write Set_Sqref;
  end;

{ IXMLSheetFormatPrType }

  IXMLSheetFormatPrType = interface(IXMLNode)
    ['{AB8BA82E-AB36-4BA9-B774-4FD875D34372}']
    { Property Accessors }
    function Get_DefaultRowHeight: WideString;
    procedure Set_DefaultRowHeight(Value: WideString);
    { Methods & Properties }
    property DefaultRowHeight: WideString read Get_DefaultRowHeight write Set_DefaultRowHeight;
  end;

{ IXMLColsType }

  IXMLColsType = interface(IXMLNodeCollection)
    ['{80185C96-EA03-42BF-B426-3C6C4D06CB7E}']
    { Property Accessors }
    function Get_Col(Index: Integer): IXMLColType;
    { Methods & Properties }
    function Add: IXMLColType;
    function Insert(const Index: Integer): IXMLColType;
    property Col[Index: Integer]: IXMLColType read Get_Col; default;
  end;

{ IXMLColType }

  IXMLColType = interface(IXMLNode)
    ['{F1AE437E-35A3-4BB2-9495-B451D4CA453B}']
    { Property Accessors }
    function Get_Min: Integer;
    function Get_Max: Integer;
    function Get_Width: WideString;
    function Get_CustomWidth: Integer;
    procedure Set_Min(Value: Integer);
    procedure Set_Max(Value: Integer);
    procedure Set_Width(Value: WideString);
    procedure Set_CustomWidth(Value: Integer);
    { Methods & Properties }
    property Min: Integer read Get_Min write Set_Min;
    property Max: Integer read Get_Max write Set_Max;
    property Width: WideString read Get_Width write Set_Width;
    property CustomWidth: Integer read Get_CustomWidth write Set_CustomWidth;
  end;

{ IXMLSheetDataType }

  IXMLSheetDataType = interface(IXMLNodeCollection)
    ['{496A0137-1800-4A59-8C02-B02B6363AAA8}']
    { Property Accessors }
    function Get_Row(Index: Integer): IXMLRowType;
    { Methods & Properties }
    function Add: IXMLRowType;
    function Insert(const Index: Integer): IXMLRowType;
    property Row[Index: Integer]: IXMLRowType read Get_Row; default;
  end;

{ IXMLRowType }

  IXMLRowType = interface(IXMLNodeCollection)
    ['{77767BF9-A947-4EBB-9D4E-2C0A986C003C}']
    { Property Accessors }
    function Get_R: Integer;
    function Get_Spans: WideString;
    function Get_C(Index: Integer): IXMLCType;
    procedure Set_R(Value: Integer);
    procedure Set_Spans(Value: WideString);
    { Methods & Properties }
    function Add: IXMLCType;
    function Insert(const Index: Integer): IXMLCType;
    property R: Integer read Get_R write Set_R;
    property Spans: WideString read Get_Spans write Set_Spans;
    property C[Index: Integer]: IXMLCType read Get_C; default;
  end;

{ IXMLCType }

  IXMLCType = interface(IXMLNode)
    ['{04191FC1-3E84-4403-827A-C02FEC75AC92}']
    { Property Accessors }
    function Get_R: WideString;
    function Get_S: Variant;
    function Get_T: Variant;
    function Get_V: Variant;
    function Get_Is_: IXMLIsType;
    function Get_F: Variant;
    procedure Set_R(Value: WideString);
    procedure Set_S(Value: Variant);
    procedure Set_T(Value: Variant);
    procedure Set_V(Value: Variant);
    procedure Set_F(Value: Variant);
    { Methods & Properties }
    property R: WideString read Get_R write Set_R;
    property S: Variant read Get_S write Set_S;
    property T: Variant read Get_T write Set_T;
    property V: Variant read Get_V write Set_V;
    property F: Variant read Get_F write Set_F;
    property Is_: IXMLIsType read Get_Is_;
  end;

{ IXMLIsType }

  IXMLIsType = interface(IXMLNode)
    ['{8B0D6E93-AAFA-456A-B0E6-7C79023E96BD}']
    { Property Accessors }
    function Get_T: WideString;
    procedure Set_T(Value: WideString);
    { Methods & Properties }
    property T: WideString read Get_T write Set_T;
  end;

{ IXMLMergeCellsType }

  IXMLMergeCellsType = interface(IXMLNodeCollection)
    ['{9FBC6C36-8E2A-4ACF-A3E5-153756DDA250}']
    { Property Accessors }
    function Get_Count: Integer;
    function Get_MergeCell(Index: Integer): IXMLMergeCellType;
    procedure Set_Count(Value: Integer);
    { Methods & Properties }
    function Add: IXMLMergeCellType;
    function Insert(const Index: Integer): IXMLMergeCellType;
    property Count: Integer read Get_Count write Set_Count;
    property MergeCell[Index: Integer]: IXMLMergeCellType read Get_MergeCell; default;
  end;

{ IXMLMergeCellType }

  IXMLMergeCellType = interface(IXMLNode)
    ['{08136422-5E3D-45CC-BEBE-6DC43071056A}']
    { Property Accessors }
    function Get_Ref: WideString;
    procedure Set_Ref(Value: WideString);
    { Methods & Properties }
    property Ref: WideString read Get_Ref write Set_Ref;
  end;

{ IXMLPageMarginsType }

  IXMLPageMarginsType = interface(IXMLNode)
    ['{6E41780E-3351-415E-9A58-F14535925441}']
    { Property Accessors }
    function Get_Left: WideString;
    function Get_Right: WideString;
    function Get_Top: WideString;
    function Get_Bottom: WideString;
    function Get_Header: WideString;
    function Get_Footer: WideString;
    procedure Set_Left(Value: WideString);
    procedure Set_Right(Value: WideString);
    procedure Set_Top(Value: WideString);
    procedure Set_Bottom(Value: WideString);
    procedure Set_Header(Value: WideString);
    procedure Set_Footer(Value: WideString);
    { Methods & Properties }
    property Left: WideString read Get_Left write Set_Left;
    property Right: WideString read Get_Right write Set_Right;
    property Top: WideString read Get_Top write Set_Top;
    property Bottom: WideString read Get_Bottom write Set_Bottom;
    property Header: WideString read Get_Header write Set_Header;
    property Footer: WideString read Get_Footer write Set_Footer;
  end;

{ IXMLPageSetupType }

  IXMLPageSetupType = interface(IXMLNode)
    ['{21C292DC-159A-4D7F-ACC7-63B3F1C74EC6}']
    { Property Accessors }
    function Get_Orientation: WideString;
    function Get_HorizontalDpi: Integer;
    function Get_VerticalDpi: Integer;
    function Get_Id: WideString;
    procedure Set_Orientation(Value: WideString);
    procedure Set_HorizontalDpi(Value: Integer);
    procedure Set_VerticalDpi(Value: Integer);
    procedure Set_Id(Value: WideString);
    { Methods & Properties }
    property Orientation: WideString read Get_Orientation write Set_Orientation;
    property HorizontalDpi: Integer read Get_HorizontalDpi write Set_HorizontalDpi;
    property VerticalDpi: Integer read Get_VerticalDpi write Set_VerticalDpi;
    property Id: WideString read Get_Id write Set_Id;
  end;

{ IXMLColType2 }

  IXMLColType2 = interface(IXMLNode)
    ['{3289C649-77E2-4FFB-B153-45E8C4C6242C}']
    { Property Accessors }
    function Get_Min: Integer;
    function Get_Max: Integer;
    function Get_Width: WideString;
    function Get_BestFit: Integer;
    function Get_CustomWidth: Integer;
    procedure Set_Min(Value: Integer);
    procedure Set_Max(Value: Integer);
    procedure Set_Width(Value: WideString);
    procedure Set_BestFit(Value: Integer);
    procedure Set_CustomWidth(Value: Integer);
    { Methods & Properties }
    property Min: Integer read Get_Min write Set_Min;
    property Max: Integer read Get_Max write Set_Max;
    property Width: WideString read Get_Width write Set_Width;
    property BestFit: Integer read Get_BestFit write Set_BestFit;
    property CustomWidth: Integer read Get_CustomWidth write Set_CustomWidth;
  end;

{ IXMLCType2 }

  IXMLCType2 = interface(IXMLNode)
    ['{BC50F887-A61D-489C-B289-CB12182D1D07}']
    { Property Accessors }
    function Get_R: WideString;
    function Get_S: Integer;
    procedure Set_R(Value: WideString);
    procedure Set_S(Value: Integer);
    { Methods & Properties }
    property R: WideString read Get_R write Set_R;
    property S: Integer read Get_S write Set_S;
  end;

{ IXMLCType22 }

  IXMLCType22 = interface(IXMLNode)
    ['{0020ABCB-598D-446B-80C9-63764C27AF32}']
    { Property Accessors }
    function Get_R: WideString;
    function Get_S: Integer;
    procedure Set_R(Value: WideString);
    procedure Set_S(Value: Integer);
    { Methods & Properties }
    property R: WideString read Get_R write Set_R;
    property S: Integer read Get_S write Set_S;
  end;

{ IXMLCType222 }

  IXMLCType222 = interface(IXMLNode)
    ['{56A49869-0858-403D-B3FF-8A9A6FBB9E05}']
    { Property Accessors }
    function Get_R: WideString;
    function Get_S: Integer;
    procedure Set_R(Value: WideString);
    procedure Set_S(Value: Integer);
    { Methods & Properties }
    property R: WideString read Get_R write Set_R;
    property S: Integer read Get_S write Set_S;
  end;

{ IXMLCType2222 }

  IXMLCType2222 = interface(IXMLNode)
    ['{097934FA-8517-4EC0-8916-F83D253F63D5}']
    { Property Accessors }
    function Get_R: WideString;
    function Get_S: Integer;
    procedure Set_R(Value: WideString);
    procedure Set_S(Value: Integer);
    { Methods & Properties }
    property R: WideString read Get_R write Set_R;
    property S: Integer read Get_S write Set_S;
  end;

{ IXMLMergeCellType2 }

  IXMLMergeCellType2 = interface(IXMLNode)
    ['{1E1D2E17-F775-4111-9E9C-750448EDDC12}']
    { Property Accessors }
    function Get_Ref: WideString;
    procedure Set_Ref(Value: WideString);
    { Methods & Properties }
    property Ref: WideString read Get_Ref write Set_Ref;
  end;

{ Forward Decls }

  TXMLWorksheetType = class;
  TXMLDimensionType = class;
  TXMLSheetViewsType = class;
  TXMLSheetViewType = class;
  TXMLSelectionType = class;
  TXMLSheetFormatPrType = class;
  TXMLColsType = class;
  TXMLColType = class;
  TXMLSheetDataType = class;
  TXMLRowType = class;
  TXMLCType = class;
  TXMLMergeCellsType = class;
  TXMLMergeCellType = class;
  TXMLPageMarginsType = class;
  TXMLPageSetupType = class;
  TXMLColType2 = class;
  TXMLCType2 = class;
  TXMLCType22 = class;
  TXMLCType222 = class;
  TXMLCType2222 = class;
  TXMLMergeCellType2 = class;

{ TXMLWorksheetType }

  TXMLWorksheetType = class(TXMLNode, IXMLWorksheetType)
  protected
    { IXMLWorksheetType }
    function Get_Xmlns: WideString;
    function Get_Dimension: IXMLDimensionType;
    function Get_SheetViews: IXMLSheetViewsType;
    function Get_SheetFormatPr: IXMLSheetFormatPrType;
    function Get_Cols: IXMLColsType;
    function Get_SheetData: IXMLSheetDataType;
    function Get_MergeCells: IXMLMergeCellsType;
    function Get_PageMargins: IXMLPageMarginsType;
    function Get_PageSetup: IXMLPageSetupType;
    procedure Set_Xmlns(Value: WideString);
  public
    procedure AfterConstruction; override;
  end;

{ TXMLDimensionType }

  TXMLDimensionType = class(TXMLNode, IXMLDimensionType)
  protected
    { IXMLDimensionType }
    function Get_Ref: WideString;
    procedure Set_Ref(Value: WideString);
  end;

{ TXMLSheetViewsType }

  TXMLSheetViewsType = class(TXMLNode, IXMLSheetViewsType)
  protected
    { IXMLSheetViewsType }
    function Get_SheetView: IXMLSheetViewType;
  public
    procedure AfterConstruction; override;
  end;

{ TXMLSheetViewType }

  TXMLSheetViewType = class(TXMLNode, IXMLSheetViewType)
  protected
    { IXMLSheetViewType }
    function Get_TabSelected: Integer;
    function Get_WorkbookViewId: Integer;
    function Get_Selection: IXMLSelectionType;
    procedure Set_TabSelected(Value: Integer);
    procedure Set_WorkbookViewId(Value: Integer);
  public
    procedure AfterConstruction; override;
  end;

{ TXMLSelectionType }

  TXMLSelectionType = class(TXMLNode, IXMLSelectionType)
  protected
    { IXMLSelectionType }
    function Get_ActiveCell: WideString;
    function Get_Sqref: WideString;
    procedure Set_ActiveCell(Value: WideString);
    procedure Set_Sqref(Value: WideString);
  end;

{ TXMLSheetFormatPrType }

  TXMLSheetFormatPrType = class(TXMLNode, IXMLSheetFormatPrType)
  protected
    { IXMLSheetFormatPrType }
    function Get_DefaultRowHeight: WideString;
    procedure Set_DefaultRowHeight(Value: WideString);
  end;

{ TXMLColsType }

  TXMLColsType = class(TXMLNodeCollection, IXMLColsType)
  protected
    { IXMLColsType }
    function Get_Col(Index: Integer): IXMLColType;
    function Add: IXMLColType;
    function Insert(const Index: Integer): IXMLColType;
  public
    procedure AfterConstruction; override;
  end;

{ TXMLColType }

  TXMLColType = class(TXMLNode, IXMLColType)
  protected
    { IXMLColType }
    function Get_Min: Integer;
    function Get_Max: Integer;
    function Get_Width: WideString;
    function Get_CustomWidth: Integer;
    procedure Set_Min(Value: Integer);
    procedure Set_Max(Value: Integer);
    procedure Set_Width(Value: WideString);
    procedure Set_CustomWidth(Value: Integer);
  end;

{ TXMLSheetDataType }

  TXMLSheetDataType = class(TXMLNodeCollection, IXMLSheetDataType)
  protected
    { IXMLSheetDataType }
    function Get_Row(Index: Integer): IXMLRowType;
    function Add: IXMLRowType;
    function Insert(const Index: Integer): IXMLRowType;
  public
    procedure AfterConstruction; override;
  end;

{ TXMLRowType }

  TXMLRowType = class(TXMLNodeCollection, IXMLRowType)
  protected
    { IXMLRowType }
    function Get_R: Integer;
    function Get_Spans: WideString;
    function Get_C(Index: Integer): IXMLCType;
    procedure Set_R(Value: Integer);
    procedure Set_Spans(Value: WideString);
    function Add: IXMLCType;
    function Insert(const Index: Integer): IXMLCType;
  public
    procedure AfterConstruction; override;
  end;

{ TXMLCType }

  TXMLCType = class(TXMLNode, IXMLCType)
  protected
    { IXMLCType }
    function Get_R: WideString;
    function Get_S: Variant;
    function Get_T: Variant;
    function Get_V: Variant;
    function Get_F: Variant;
    function Get_Is_: IXMLIsType;
    procedure Set_R(Value: WideString);
    procedure Set_S(Value: Variant);
    procedure Set_T(Value: Variant);
    procedure Set_V(Value: Variant);
    procedure Set_F(Value: Variant);
  public
    procedure AfterConstruction; override;
  end;

{ TXMLIsType }

  TXMLIsType = class(TXMLNode, IXMLIsType)
  protected
    { IXMLIsType }
    function Get_T: WideString;
    procedure Set_T(Value: WideString);
  end;

{ TXMLMergeCellsType }

  TXMLMergeCellsType = class(TXMLNodeCollection, IXMLMergeCellsType)
  protected
    { IXMLMergeCellsType }
    function Get_Count: Integer;
    function Get_MergeCell(Index: Integer): IXMLMergeCellType;
    procedure Set_Count(Value: Integer);
    function Add: IXMLMergeCellType;
    function Insert(const Index: Integer): IXMLMergeCellType;
  public
    procedure AfterConstruction; override;
  end;

{ TXMLMergeCellType }

  TXMLMergeCellType = class(TXMLNode, IXMLMergeCellType)
  protected
    { IXMLMergeCellType }
    function Get_Ref: WideString;
    procedure Set_Ref(Value: WideString);
  end;

{ TXMLPageMarginsType }

  TXMLPageMarginsType = class(TXMLNode, IXMLPageMarginsType)
  protected
    { IXMLPageMarginsType }
    function Get_Left: WideString;
    function Get_Right: WideString;
    function Get_Top: WideString;
    function Get_Bottom: WideString;
    function Get_Header: WideString;
    function Get_Footer: WideString;
    procedure Set_Left(Value: WideString);
    procedure Set_Right(Value: WideString);
    procedure Set_Top(Value: WideString);
    procedure Set_Bottom(Value: WideString);
    procedure Set_Header(Value: WideString);
    procedure Set_Footer(Value: WideString);
  end;

{ TXMLPageSetupType }

  TXMLPageSetupType = class(TXMLNode, IXMLPageSetupType)
  protected
    { IXMLPageSetupType }
    function Get_Orientation: WideString;
    function Get_HorizontalDpi: Integer;
    function Get_VerticalDpi: Integer;
    function Get_Id: WideString;
    procedure Set_Orientation(Value: WideString);
    procedure Set_HorizontalDpi(Value: Integer);
    procedure Set_VerticalDpi(Value: Integer);
    procedure Set_Id(Value: WideString);
  end;

{ TXMLColType2 }

  TXMLColType2 = class(TXMLNode, IXMLColType2)
  protected
    { IXMLColType2 }
    function Get_Min: Integer;
    function Get_Max: Integer;
    function Get_Width: WideString;
    function Get_BestFit: Integer;
    function Get_CustomWidth: Integer;
    procedure Set_Min(Value: Integer);
    procedure Set_Max(Value: Integer);
    procedure Set_Width(Value: WideString);
    procedure Set_BestFit(Value: Integer);
    procedure Set_CustomWidth(Value: Integer);
  end;

{ TXMLCType2 }

  TXMLCType2 = class(TXMLNode, IXMLCType2)
  protected
    { IXMLCType2 }
    function Get_R: WideString;
    function Get_S: Integer;
    procedure Set_R(Value: WideString);
    procedure Set_S(Value: Integer);
  end;

{ TXMLCType22 }

  TXMLCType22 = class(TXMLNode, IXMLCType22)
  protected
    { IXMLCType22 }
    function Get_R: WideString;
    function Get_S: Integer;
    procedure Set_R(Value: WideString);
    procedure Set_S(Value: Integer);
  end;

{ TXMLCType222 }

  TXMLCType222 = class(TXMLNode, IXMLCType222)
  protected
    { IXMLCType222 }
    function Get_R: WideString;
    function Get_S: Integer;
    procedure Set_R(Value: WideString);
    procedure Set_S(Value: Integer);
  end;

{ TXMLCType2222 }

  TXMLCType2222 = class(TXMLNode, IXMLCType2222)
  protected
    { IXMLCType2222 }
    function Get_R: WideString;
    function Get_S: Integer;
    procedure Set_R(Value: WideString);
    procedure Set_S(Value: Integer);
  end;

{ TXMLMergeCellType2 }

  TXMLMergeCellType2 = class(TXMLNode, IXMLMergeCellType2)
  protected
    { IXMLMergeCellType2 }
    function Get_Ref: WideString;
    procedure Set_Ref(Value: WideString);
  end;

{ Global Functions }

function Getworksheet(Doc: IXMLDocument): IXMLWorksheetType;
function Loadworksheet(const FileName: WideString): IXMLWorksheetType;
function Newworksheet: IXMLWorksheetType;

const
  TargetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
{$ENDIF}
{$ENDIF}

implementation
{$IFDEF XLSX}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    System.Variants;
  {$ELSE}
    Variants;
  {$ENDIF}

{ Global Functions }

function Getworksheet(Doc: IXMLDocument): IXMLWorksheetType;
begin
  Result := Doc.GetDocBinding('worksheet', TXMLWorksheetType, TargetNamespace) as IXMLWorksheetType;
end;

function Loadworksheet(const FileName: WideString): IXMLWorksheetType;
begin
  Result := LoadXMLDocument(FileName).GetDocBinding('worksheet', TXMLWorksheetType, TargetNamespace) as IXMLWorksheetType;
end;

function Newworksheet: IXMLWorksheetType;
begin
  Result := NewXMLDocument.GetDocBinding('worksheet', TXMLWorksheetType, TargetNamespace) as IXMLWorksheetType;
end;

{ TXMLWorksheetType }

procedure TXMLWorksheetType.AfterConstruction;
begin
  RegisterChildNode('dimension', TXMLDimensionType);
  RegisterChildNode('sheetViews', TXMLSheetViewsType);
  RegisterChildNode('sheetFormatPr', TXMLSheetFormatPrType);
  RegisterChildNode('cols', TXMLColsType);
  RegisterChildNode('sheetData', TXMLSheetDataType);
  RegisterChildNode('mergeCells', TXMLMergeCellsType);
  RegisterChildNode('pageMargins', TXMLPageMarginsType);
  RegisterChildNode('pageSetup', TXMLPageSetupType);
  inherited;
end;

function TXMLWorksheetType.Get_Xmlns: WideString;
begin
  Result := AttributeNodes['xmlns'].Text;
end;

procedure TXMLWorksheetType.Set_Xmlns(Value: WideString);
begin
  SetAttribute('xmlns', Value);
end;

function TXMLWorksheetType.Get_Dimension: IXMLDimensionType;
begin
  Result := ChildNodes['dimension'] as IXMLDimensionType;
end;

function TXMLWorksheetType.Get_SheetViews: IXMLSheetViewsType;
begin
  Result := ChildNodes['sheetViews'] as IXMLSheetViewsType;
end;

function TXMLWorksheetType.Get_SheetFormatPr: IXMLSheetFormatPrType;
begin
  Result := ChildNodes['sheetFormatPr'] as IXMLSheetFormatPrType;
end;

function TXMLWorksheetType.Get_Cols: IXMLColsType;
begin
  Result := ChildNodes['cols'] as IXMLColsType;
end;

function TXMLWorksheetType.Get_SheetData: IXMLSheetDataType;
begin
  Result := ChildNodes['sheetData'] as IXMLSheetDataType;
end;

function TXMLWorksheetType.Get_MergeCells: IXMLMergeCellsType;
begin
  Result := ChildNodes['mergeCells'] as IXMLMergeCellsType;
end;

function TXMLWorksheetType.Get_PageMargins: IXMLPageMarginsType;
begin
  Result := ChildNodes['pageMargins'] as IXMLPageMarginsType;
end;

function TXMLWorksheetType.Get_PageSetup: IXMLPageSetupType;
begin
  Result := ChildNodes['pageSetup'] as IXMLPageSetupType;
end;

{ TXMLDimensionType }

function TXMLDimensionType.Get_Ref: WideString;
begin
  Result := AttributeNodes['ref'].Text;
end;

procedure TXMLDimensionType.Set_Ref(Value: WideString);
begin
  SetAttribute('ref', Value);
end;

{ TXMLSheetViewsType }

procedure TXMLSheetViewsType.AfterConstruction;
begin
  RegisterChildNode('sheetView', TXMLSheetViewType);
  inherited;
end;

function TXMLSheetViewsType.Get_SheetView: IXMLSheetViewType;
begin
  Result := ChildNodes['sheetView'] as IXMLSheetViewType;
end;

{ TXMLSheetViewType }

procedure TXMLSheetViewType.AfterConstruction;
begin
  RegisterChildNode('selection', TXMLSelectionType);
  inherited;
end;

function TXMLSheetViewType.Get_TabSelected: Integer;
begin
  Result := AttributeNodes['tabSelected'].NodeValue;
end;

procedure TXMLSheetViewType.Set_TabSelected(Value: Integer);
begin
  SetAttribute('tabSelected', Value);
end;

function TXMLSheetViewType.Get_WorkbookViewId: Integer;
begin
  Result := AttributeNodes['workbookViewId'].NodeValue;
end;

procedure TXMLSheetViewType.Set_WorkbookViewId(Value: Integer);
begin
  SetAttribute('workbookViewId', Value);
end;

function TXMLSheetViewType.Get_Selection: IXMLSelectionType;
begin
  Result := ChildNodes['selection'] as IXMLSelectionType;
end;

{ TXMLSelectionType }

function TXMLSelectionType.Get_ActiveCell: WideString;
begin
  Result := AttributeNodes['activeCell'].Text;
end;

procedure TXMLSelectionType.Set_ActiveCell(Value: WideString);
begin
  SetAttribute('activeCell', Value);
end;

function TXMLSelectionType.Get_Sqref: WideString;
begin
  Result := AttributeNodes['sqref'].Text;
end;

procedure TXMLSelectionType.Set_Sqref(Value: WideString);
begin
  SetAttribute('sqref', Value);
end;

{ TXMLSheetFormatPrType }

function TXMLSheetFormatPrType.Get_DefaultRowHeight: WideString;
begin
  Result := AttributeNodes['defaultRowHeight'].Text;
end;

procedure TXMLSheetFormatPrType.Set_DefaultRowHeight(Value: WideString);
begin
  SetAttribute('defaultRowHeight', Value);
end;

{ TXMLColsType }

procedure TXMLColsType.AfterConstruction;
begin
  RegisterChildNode('col', TXMLColType);
  ItemTag := 'col';
  ItemInterface := IXMLColType;
  inherited;
end;

function TXMLColsType.Get_Col(Index: Integer): IXMLColType;
begin
  Result := List[Index] as IXMLColType;
end;

function TXMLColsType.Add: IXMLColType;
begin
  Result := AddItem(-1) as IXMLColType;
end;

function TXMLColsType.Insert(const Index: Integer): IXMLColType;
begin
  Result := AddItem(Index) as IXMLColType;
end;

{ TXMLColType }

function TXMLColType.Get_Min: Integer;
begin
  Result := AttributeNodes['min'].NodeValue;
end;

procedure TXMLColType.Set_Min(Value: Integer);
begin
  SetAttribute('min', Value);
end;

function TXMLColType.Get_Max: Integer;
begin
  Result := AttributeNodes['max'].NodeValue;
end;

procedure TXMLColType.Set_Max(Value: Integer);
begin
  SetAttribute('max', Value);
end;

function TXMLColType.Get_Width: WideString;
begin
  Result := AttributeNodes['width'].Text;
end;

procedure TXMLColType.Set_Width(Value: WideString);
begin
  SetAttribute('width', Value);
end;

function TXMLColType.Get_CustomWidth: Integer;
begin
  Result := AttributeNodes['customWidth'].NodeValue;
end;

procedure TXMLColType.Set_CustomWidth(Value: Integer);
begin
  SetAttribute('customWidth', Value);
end;

{ TXMLSheetDataType }

procedure TXMLSheetDataType.AfterConstruction;
begin
  RegisterChildNode('row', TXMLRowType);
  ItemTag := 'row';
  ItemInterface := IXMLRowType;
  inherited;
end;

function TXMLSheetDataType.Get_Row(Index: Integer): IXMLRowType;
begin
  Result := List[Index] as IXMLRowType;
end;

function TXMLSheetDataType.Add: IXMLRowType;
begin
  Result := AddItem(-1) as IXMLRowType;
end;

function TXMLSheetDataType.Insert(const Index: Integer): IXMLRowType;
begin
  Result := AddItem(Index) as IXMLRowType;
end;

{ TXMLRowType }

procedure TXMLRowType.AfterConstruction;
begin
  RegisterChildNode('c', TXMLCType);
  ItemTag := 'c';
  ItemInterface := IXMLCType;
  inherited;
end;

function TXMLRowType.Get_R: Integer;
begin
  Result := AttributeNodes[WideString('r')].NodeValue;
end;

procedure TXMLRowType.Set_R(Value: Integer);
begin
  SetAttribute(WideString('r'), Value);
end;

function TXMLRowType.Get_Spans: WideString;
begin
  Result := AttributeNodes['spans'].Text;
end;

procedure TXMLRowType.Set_Spans(Value: WideString);
begin
  SetAttribute('spans', Value);
end;

function TXMLRowType.Get_C(Index: Integer): IXMLCType;
begin
  Result := List[Index] as IXMLCType;
end;

function TXMLRowType.Add: IXMLCType;
begin
  Result := AddItem(-1) as IXMLCType;
end;

function TXMLRowType.Insert(const Index: Integer): IXMLCType;
begin
  Result := AddItem(Index) as IXMLCType;
end;

{ TXMLCType }

procedure TXMLCType.AfterConstruction;
begin
  RegisterChildNode('is', TXMLIsType);
  inherited;
end;

function TXMLCType.Get_R: WideString;
begin
  Result := AttributeNodes[WideString('r')].Text;
end;

procedure TXMLCType.Set_R(Value: WideString);
begin
  SetAttribute(WideString('r'), Value);
end;

function TXMLCType.Get_S: Variant;
begin
  Result := AttributeNodes[WideString('s')].NodeValue;
end;

procedure TXMLCType.Set_S(Value: Variant);
begin
  SetAttribute(WideString('s'), Value);
end;

function TXMLCType.Get_T: Variant;
begin
  Result := AttributeNodes[WideString('t')].Text;
end;

procedure TXMLCType.Set_T(Value: Variant);
begin
  SetAttribute(WideString('t'), Value);
end;

function TXMLCType.Get_V: Variant;
begin
  Result := ChildNodes[WideString('v')].NodeValue;
end;

procedure TXMLCType.Set_V(Value: Variant);
begin
  ChildNodes[WideString('v')].NodeValue := Value;
end;

function TXMLCType.Get_F: Variant;
begin
  Result := ChildNodes[WideString('f')].Text;
end;

procedure TXMLCType.Set_F(Value: Variant);
begin
  ChildNodes[WideString('f')].NodeValue := Value;
end;

function TXMLCType.Get_Is_: IXMLIsType;
begin
  Result := ChildNodes['is'] as IXMLIsType;
end;

{ TXMLIsType }

function TXMLIsType.Get_T: WideString;
begin
  Result := ChildNodes[WideString('t')].Text;
end;

procedure TXMLIsType.Set_T(Value: WideString);
begin
  ChildNodes[WideString('t')].NodeValue := Value;
end;

{ TXMLMergeCellsType }

procedure TXMLMergeCellsType.AfterConstruction;
begin
  RegisterChildNode('mergeCell', TXMLMergeCellType);
  ItemTag := 'mergeCell';
  ItemInterface := IXMLMergeCellType;
  inherited;
end;

function TXMLMergeCellsType.Get_Count: Integer;
  var
    CountValue: Variant;
begin
  CountValue := AttributeNodes['count'].NodeValue;

  if VarIsNull(CountValue) then
    Result := 0
  else
    Result := CountValue;
end;

procedure TXMLMergeCellsType.Set_Count(Value: Integer);
begin
  SetAttribute('count', Value);
end;

function TXMLMergeCellsType.Get_MergeCell(Index: Integer): IXMLMergeCellType;
begin
  Result := List[Index] as IXMLMergeCellType;
end;

function TXMLMergeCellsType.Add: IXMLMergeCellType;
begin
  Result := AddItem(-1) as IXMLMergeCellType;
end;

function TXMLMergeCellsType.Insert(const Index: Integer): IXMLMergeCellType;
begin
  Result := AddItem(Index) as IXMLMergeCellType;
end;

{ TXMLMergeCellType }

function TXMLMergeCellType.Get_Ref: WideString;
begin
  Result := AttributeNodes['ref'].Text;
end;

procedure TXMLMergeCellType.Set_Ref(Value: WideString);
begin
  SetAttribute('ref', Value);
end;

{ TXMLPageMarginsType }

function TXMLPageMarginsType.Get_Left: WideString;
begin
  Result := AttributeNodes['left'].Text;
end;

procedure TXMLPageMarginsType.Set_Left(Value: WideString);
begin
  SetAttribute('left', Value);
end;

function TXMLPageMarginsType.Get_Right: WideString;
begin
  Result := AttributeNodes['right'].Text;
end;

procedure TXMLPageMarginsType.Set_Right(Value: WideString);
begin
  SetAttribute('right', Value);
end;

function TXMLPageMarginsType.Get_Top: WideString;
begin
  Result := AttributeNodes['top'].Text;
end;

procedure TXMLPageMarginsType.Set_Top(Value: WideString);
begin
  SetAttribute('top', Value);
end;

function TXMLPageMarginsType.Get_Bottom: WideString;
begin
  Result := AttributeNodes['bottom'].Text;
end;

procedure TXMLPageMarginsType.Set_Bottom(Value: WideString);
begin
  SetAttribute('bottom', Value);
end;

function TXMLPageMarginsType.Get_Header: WideString;
begin
  Result := AttributeNodes['header'].Text;
end;

procedure TXMLPageMarginsType.Set_Header(Value: WideString);
begin
  SetAttribute('header', Value);
end;

function TXMLPageMarginsType.Get_Footer: WideString;
begin
  Result := AttributeNodes['footer'].Text;
end;

procedure TXMLPageMarginsType.Set_Footer(Value: WideString);
begin
  SetAttribute('footer', Value);
end;

{ TXMLPageSetupType }

function TXMLPageSetupType.Get_Orientation: WideString;
begin
  Result := AttributeNodes['orientation'].Text;
end;

procedure TXMLPageSetupType.Set_Orientation(Value: WideString);
begin
  SetAttribute('orientation', Value);
end;

function TXMLPageSetupType.Get_HorizontalDpi: Integer;
begin
  Result := AttributeNodes['horizontalDpi'].NodeValue;
end;

procedure TXMLPageSetupType.Set_HorizontalDpi(Value: Integer);
begin
  SetAttribute('horizontalDpi', Value);
end;

function TXMLPageSetupType.Get_VerticalDpi: Integer;
begin
  Result := AttributeNodes['verticalDpi'].NodeValue;
end;

procedure TXMLPageSetupType.Set_VerticalDpi(Value: Integer);
begin
  SetAttribute('verticalDpi', Value);
end;

function TXMLPageSetupType.Get_Id: WideString;
begin
  Result := AttributeNodes['id'].Text;
end;

procedure TXMLPageSetupType.Set_Id(Value: WideString);
begin
  SetAttribute('id', Value);
end;

{ TXMLColType2 }

function TXMLColType2.Get_Min: Integer;
begin
  Result := AttributeNodes['min'].NodeValue;
end;

procedure TXMLColType2.Set_Min(Value: Integer);
begin
  SetAttribute('min', Value);
end;

function TXMLColType2.Get_Max: Integer;
begin
  Result := AttributeNodes['max'].NodeValue;
end;

procedure TXMLColType2.Set_Max(Value: Integer);
begin
  SetAttribute('max', Value);
end;

function TXMLColType2.Get_Width: WideString;
begin
  Result := AttributeNodes['width'].Text;
end;

procedure TXMLColType2.Set_Width(Value: WideString);
begin
  SetAttribute('width', Value);
end;

function TXMLColType2.Get_BestFit: Integer;
begin
  Result := AttributeNodes['bestFit'].NodeValue;
end;

procedure TXMLColType2.Set_BestFit(Value: Integer);
begin
  SetAttribute('bestFit', Value);
end;

function TXMLColType2.Get_CustomWidth: Integer;
begin
  Result := AttributeNodes['customWidth'].NodeValue;
end;

procedure TXMLColType2.Set_CustomWidth(Value: Integer);
begin
  SetAttribute('customWidth', Value);
end;

{ TXMLCType2 }

function TXMLCType2.Get_R: WideString;
begin
  Result := AttributeNodes[WideString('r')].Text;
end;

procedure TXMLCType2.Set_R(Value: WideString);
begin
  SetAttribute(WideString('r'), Value);
end;

function TXMLCType2.Get_S: Integer;
begin
  Result := AttributeNodes[WideString('s')].NodeValue;
end;

procedure TXMLCType2.Set_S(Value: Integer);
begin
  SetAttribute(WideString('s'), Value);
end;

{ TXMLCType22 }

function TXMLCType22.Get_R: WideString;
begin
  Result := AttributeNodes[WideString('r')].Text;
end;

procedure TXMLCType22.Set_R(Value: WideString);
begin
  SetAttribute(WideString('r'), Value);
end;

function TXMLCType22.Get_S: Integer;
begin
  Result := AttributeNodes[WideString('s')].NodeValue;
end;

procedure TXMLCType22.Set_S(Value: Integer);
begin
  SetAttribute(WideString('s'), Value);
end;

{ TXMLCType222 }

function TXMLCType222.Get_R: WideString;
begin
  Result := AttributeNodes[WideString('r')].Text;
end;

procedure TXMLCType222.Set_R(Value: WideString);
begin
  SetAttribute(WideString('r'), Value);
end;

function TXMLCType222.Get_S: Integer;
begin
  Result := AttributeNodes[WideString('s')].NodeValue;
end;

procedure TXMLCType222.Set_S(Value: Integer);
begin
  SetAttribute(WideString('s'), Value);
end;

{ TXMLCType2222 }

function TXMLCType2222.Get_R: WideString;
begin
  Result := AttributeNodes[WideString('r')].Text;
end;

procedure TXMLCType2222.Set_R(Value: WideString);
begin
  SetAttribute(WideString('r'), Value);
end;

function TXMLCType2222.Get_S: Integer;
begin
  Result := AttributeNodes[WideString('s')].NodeValue;
end;

procedure TXMLCType2222.Set_S(Value: Integer);
begin
  SetAttribute(WideString('s'), Value);
end;

{ TXMLMergeCellType2 }

function TXMLMergeCellType2.Get_Ref: WideString;
begin
  Result := AttributeNodes['ref'].Text;
end;

procedure TXMLMergeCellType2.Set_Ref(Value: WideString);
begin
  SetAttribute('ref', Value);
end;
{$ENDIF}
{$ENDIF}


end.