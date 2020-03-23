unit QImport3Consts;

interface

resourcestring
  // Errors
  QIE_NoDataSet            = 'DataSet is not defined';
  QIE_NoDBGrid             = 'DBGrid is not defined'; // since 1.7
  QIE_NoListView           = 'ListView is not defined'; // since 1.7
  QIE_NoStringGrid         = 'StrinGrid is not defined'; // since 1.7
  QIE_NoSource             = 'Source DataSet is not defined'; // since 1.5
  QIE_NoFileName           = 'File name is not specified';
  QIE_FileNotExists        = 'File %s does not exist';
  QIE_MappingEmpty         = 'Map is empty';
  QIE_CanceledByUser       = 'Import canceled by user';
  QIE_MapMissing           = 'Missing definition of map';
  QIE_FieldNotFound        = 'Column %s is defined in the MAP is not found in the import destination';
  QIE_SourceFieldNotFound  = 'Field %s not found in the source DataSet'; // since 1.5
  QIE_DataSetClosed        = 'Cannot perform this operation on a closed Import Destination';
  QIE_AllowedImportsEmpty  = 'Property AllowedImports can not contain empty set';
  QIE_KeyColumnsNotDefined = 'Key columns are not defined'; // since 1.7
  QIE_KeyColumnNotFound    = 'Key column %s is not found in the import destination'; // since 1.7
  QIE_XMLFieldListEmpty    = 'XML field list is empty'; // since 2.0
  QIE_NeedDefineTextToFind = 'Text to find must be defined'; // since 2.0
  QIE_GridCaptionPos       = 'GridCaptionRow cannot be greater than GridStartRow';
  QIE_XlsxMapError         = 'Error while parsing map';

  // Filters
  QIF_TXT      = 'Text files (*.txt)|*.txt|All files (*.*)|*.*';
  QIF_ASCII    = 'ASCII files (*.txt;*.csv)|*.txt;*.csv|All files (*.*)|*.*';
  QIF_CSV      = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*';
  QIF_DBF      = 'DBF files (*.dbf)|*.dbf|All files (*.*)|*.*';
  QIF_HTML     = 'HTML files (*.htm, *.html)|*.htm; *.html;|All files (*.*)|*.*'; 
  QIF_XML      = 'XML files (*.xml)|*.xml|All files (*.*)|*.*'; // since 2.0
  QIF_XLSX     = 'Microsoft  Excel 2007 files (*.xlsx)|*.xlsx|All files (*.*)|*.*';
  QIF_DOCX     = 'Microsoft Word 2007 files (*.docx)|*.docx|All files (*.*)|*.*';
  QIF_ODS      = 'OpenDocument Spreadsheet files(*.ods)|*.ods|All files (*.*)|*.*';
  QIF_ODT      = 'OpenDocument Text files(*.odt)|*.odt|All files (*.*)|*.*';
  QIF_XLS      = 'Microsoft Excel files (*.xls)|*.xls|All files (*.*)|*.*';
  QIF_Access   = 'Microsoft Access files (*.mdb)|*.mdb|All files (*.*)|*.*'; // since 1.4
  QIF_ALL      = 'All allowed files (*.txt;*.csv;*.dbf;*.xls)|*.txt;*.csv;*.dbf;*.xls';
  QIF_TEMPLATE = 'QuickImport Template files (*.imp)| *.imp'; // since 1.2

  // Loading
  QIL_Loading                = 'Loading...';                       // since 2.0

  // Defaults
  QID_BooleanTrue            = 'True';
  QID_BooleanFalse           = 'False';
  QID_NullValue              = 'Null';                             // since 1.8

  // Wizard
  QIW_importWizard           = 'Import data wizard';
  QIW_ImportFrom             = 'Import from';
  QIW_XLS                    = 'MS Excel';
  QIW_Access                 = 'MS Access';                        // since 1.4
  QIW_DBF                    = 'DBF';
  QIW_HTML                   = 'HTML';
  QIW_XML                    = 'XML';                              // since 2.0
  QIW_XMLDoc                 = 'XML Document';
  QIW_TXT                    = 'Text file';
  QIW_CSV                    = 'CSV file';
  QIW_Comma                  = 'Comma';                            // since 2.0
  QIW_Quote                  = 'Quote';                            // since 1.65
  QIW_FileName               = 'Source file name';
  QIW_NeedCancelCaption      = 'Are you sure...';
  QIW_NeedCancel             = 'Are you sure you want to abort import?';
  QIW_Question               = 'Question';                         // since 2.0
  QIW_CancelWizard           = 'Do you want to quit?';             // since 2.0

  QIW_Help                   = '&Help';
  QIW_Back                   = '< &Back';
  QIW_Next                   = '&Next >';
  QIW_Cancel                 = '&Cancel';
  QIW_Execute                = '&Execute';
  QIW_Step                   = 'Step %d of %d';
  QIW_TemplateOptions        = ' Template Options ';               // since 2.0
  QIW_GoToLastPage           = 'Go to the last page after '        + 
                               'loading template';                 // since 2.0
  QIW_AutoSaveTemplate       = 'Save template automatically when ' +
                               'wizard closes';                    // since 2.0
  QIW_LoadTemplate           = 'Load template from file...';       // since 1.2
  QIW_SaveTemplate           = 'Save template to file...';         // since 1.2

  QIW_CommitOptions          = 'Commit Options';                   // since 1.7
  QIW_Commit                 = 'Commit';
  QIW_CommitAfterDone        = 'Commit when done';
  QIW_CommitRecCount         = 'Commit after each';
  QIW_Records                = 'record(s)';
  QIW_RecordCount            = 'Record count';
  QIW_ImportAllRecords       = 'Import all records';
  QIW_ImportRecCount         = 'Import only first';
  QIW_CloseAfterImport       = 'Close wizard after import';        // since 1.8

  QIW_ImportAdvanced         = 'Advanced';                         // since 1.7
  QIW_AddType                = 'Add Type';                         // since 1.6
  QIW_AddType_Append         = 'Append';                           // since 1.6
  QIW_AddType_Insert         = 'Insert';                           // since 1.6

  QIWIM_Caption              = ' Import Mode ';                    // since 1.7
  QIWIM_Insert_All           = 'Insert All';                       // since 1.7
  QIWIM_Insert_New           = 'Insert New';                       // since 1.7
  QIWIM_Update               = 'Update';                           // since 1.7
  QIWIM_Update_or_Insert     = 'Update or Insert';                 // since 1.7
  QIWIM_Delete               = 'Delete';                           // since 1.7
  QIWIM_Delete_or_Insert     = 'Delete or Insert';                 // since 1.7

  QIW_AvailableColumns       = 'Available Columns';                // since 1.7
  QIW_SelectedColumns        = 'Key Columns';                      // since 1.7

  QIW_ErrorLog               = 'Error log';
  QIW_EnableErrorLog         = 'Enable error log';
  QIW_ErrorLogFileName       = 'Error Log File Name';              // since 1.8
  QIW_RewriteErrorLogFile    = 'Rewrite error log file if it '     +
                               'exists';                           // since 1.8
  QIW_ShowErrorLog           = 'Show error log after import';
  QIW_ErrorLogStarted        = '%s - Import process started';      // since 1.8
  QIW_ErrorLogFinished       = '%s - Import process finished';     // since 1.8
  QIW_SomeErrorsFound        = '%d error(s) found';                // since 1.8
  QIW_NoErrorsFound          = 'No error found';                   // since 1.8

  QIW_BaseFormats            = 'Base formats';
  QIW_DateTimeFormats        = 'Date and Time formats';            // since 1.65
  QIW_Separators             = 'Separators';                       // since 1.65
  QIW_DecimalSeparator       = 'Decimal';
  QIW_ThousandSeparator      = 'Thousand';
  QIW_ShortDateFormat        = 'Short date';
  QIW_LongDateFormat         = 'Long date';
  QIW_DateSeparator          = 'Date';                             // since 1.65
  QIW_ShortTimeFormat        = 'Short time';
  QIW_LongTimeFormat         = 'Long time';
  QIW_TimeSeparator          = 'Time';                             // since 1.65

  QIW_BooleanTrue            = 'Boolean true';
  QIW_BooleanFalse           = 'Boolean false';
  QIW_NullValue              = 'Null Values';                      // since 2.0

  QIWDF_Caption              = 'Data formats';
  QIWDF_Fields               = 'Fields';
  QIWDF_Tuning               = 'Tuning';                           // since 1.6
  QIWDF_Replacements         = 'Replacements';                     // since 1.6
  QIWDF_TextToFind           = 'Text to find';                     // since 2.0
  QIWDF_ReplaceWith          = 'Replace with';                     // since 2.0
  QIWDF_IgnoreCase           = 'Ignore case';                      // since 2.0
  QIWDF_IgnoreCase_Yes       = 'Yes';                              // since 2.0
  QIWDF_IgnoreCase_No        = 'No';                               // since 2.0
  QIWDF_AddReplacement       = 'Add Replacement';                  // since 2.0
  QIWDF_EditReplacement      = 'Edit Replacement';                 // since 2.0
  QIWDF_DelReplacement       = 'Del Replacement';                  // since 2.0
  QIWDF_GeneratorValue       = 'Generator value';
  QIWDF_GeneratorStep        = 'Generator step';
  QIWDF_ConstantValue        = 'Constant value';
  QIWDF_NullValue            = 'Null value';
  QIWDF_DefaultValue         = 'Default value';
  QIWDF_LeftQuote            = 'Left quote';
  QIWDF_RightQuote           = 'Right quote';
  QIWDF_QuoteAction          = 'Quote action';
  QIWDF_QuoteNone            = 'None';
  QIWDF_QuoteAdd             = 'Add';
  QIWDF_QuoteRemove          = 'Remove';
  QIWDF_CharCase             = 'Char case';
  QIWDF_CharCaseNone         = 'As Is';
  QIWDF_CharCaseUpper        = 'Upper';
  QIWDF_CharCaseLower        = 'Lower';
  QIWDF_CharCaseUpperFirst   = 'Upper first';
  QIWDF_CharCaseUpperWord    = 'Upper first word';
  QIWDF_CharSet              = 'Char set';
  QIWDF_CharSetNone          = 'As Is';
  QIWDF_CharSetAnsi          = 'Ansi';
  QIWDF_CharSetOem           = 'Oem';

  QIWDF_RE_Title             = 'Replacement';                      // since 2.0
  QIWDF_RE_TextToFind        = 'Text to find';                     // since 2.0
  QIWDF_RE_ReplaceWith       = 'Replace with';                     // since 2.0
  QIWDF_RE_IgnoreCase        = 'Ignore case';                      // since 2.0

  QIW_TXT_Fields             = 'Fields';
  QIW_TXT_Fields_Pos         = 'P';                                // since 2.0
  QIW_TXT_Fields_Size        = 'S';                                // since 2.0
  QIW_TXT_Clear              = 'Clear';
  QIW_TXT_SkipLines          = 'Skip line(s)';
  QIW_TXT_Tip                = 'Double click to add or remove '    +
                               'column separators. Click at the '  +
                               'area between separators to '       +
                               'define the imported column.';      // since 2.0
  QIW_TXT_Encoding           = 'Encoding';  // dee

  QIW_CSV_Fields             = 'Fields';
  QIW_CSV_ColNumber          = 'Col Number';                       // since 2.0
  QIW_CSV_SkipLines          = 'Skip line(s)';
  QIW_CSV_AutoFill           = 'Auto fill';                        // since 1.2
  QIW_CSV_Clear              = 'Clear';                            // since 1.2
  QIW_CSV_GridCol            = 'Column_%d';                        // since 1.4
  QIW_CSV_Tip                = 'Select field name from the list '  +
                               'box, then click at the column '    +
                               'to import this field to.';

  QIW_DBF_Add                = 'Add';
  QIW_DBF_Remove             = 'Remove';
  QIW_DBF_AutoFill           = 'Auto fill';                        // since 1.2
  QIW_DBF_Clear              = 'Clear';                            // since 1.2
  QIW_DBF_SkipDeleted        = 'Skip Deleted Rows';                // since 1.7
  QIW_DBF_Tip                = 'Click the "Add" button to set '    +
                               'the accordance between the '       +
                               'imported column and the table '    +
                               'field or click the "Remove" '      +
                               'button to remove one.';

  QIW_HTML_Fields            = 'Fields';
  QIW_HTML_ColNumber         = 'Col Number';
  QIW_HTML_TableNumber       = 'Table Number';
  QIW_HTML_SkipLines         = 'Skip line(s)';
  QIW_HTML_AutoFill          = 'Auto fill';
  QIW_HTML_Clear             = 'Clear';
  QIW_HTML_GridCol           = 'Column_%d';
  QIW_HTML_Tip                = 'Select field name from the list '  +
                               'box, then click at the column '    +
                               'to import this field to.';


  QIW_XML_Add                = 'Add';                              // since 2.0
  QIW_XML_Remove             = 'Remove';                           // since 2.0
  QIW_XML_AutoFill           = 'Auto fill';                        // since 2.0
  QIW_XML_Clear              = 'Clear';                            // since 2.0
  QIW_XML_WriteOnFly         = 'Write on Fly';                     // since 2.1
  QIW_XML_Tip                = 'Click the "Add" button to set '    +
                               'the accordance between the '       +
                               'imported column and the table '    +
                               'field or click the "Remove"'       +
                               'button to remove one.';            // since 2.0

  QIW_XLS_Fields             = 'Fields';
  QIW_XLS_Ranges             = 'Ranges';                           // since 2.0
  QIW_XLS_AddRange           = 'Add Range';                        // since 2.0
  QIW_XLS_EditRange          = 'Edit Range';                       // since 2.0
  QIW_XLS_DelRange           = 'Del Range';                        // since 2.0
  QIW_XLS_MoveRangeUp        = 'Move Range Up';                    // since 2.0
  QIW_XLS_MoveRangeDown      = 'Move Range Down';                  // since 2.0
  QIW_XLS_SkipCols           = 'Skip col(s)';                      // since 2.0
  QIW_XLS_SkipRows           = 'Skip row(s)';                      // since 2.0
  QIW_XLS_AutoFillCols       = 'Auto fill cols';                   // since 1.2
  QIW_XLS_AutoFillRows       = 'Auto fill rows';                   // since 1.2
  QIW_XLS_ClearFieldRanges   = 'Clear field ranges';               // since 2.0
  QIW_XLS_ClearAllRanges     = 'Clear all ranges';                 // since 2.0
  QIW_XLS_Tip                = 'Use the Range Editor to add/edit ' +
                               'ranges or click individual cells ' +
                               'with Shift or Ctrl pressed. '      +
                               'Press Enter to apply selection '   +
                               'or Escape to cancel.';

  QIRE_Caption               = 'Range';                            // since 2.0
  QIRE_RangeType             = ' Range Type ';                     // since 2.0
  QIRE_RangeType_Col         = 'Col';                              // since 2.0
  QIRE_RangeType_Row         = 'Row';                              // since 2.0
  QIRE_RangeType_Cell        = 'Cell';                             // since 2.0
  QIRE_Start                 = ' Start ';                          // since 2.0
  QIRE_Start_DataStarted     = 'Where data started';               // since 2.0
  QIRE_Start_Col             = 'Start Col';                        // since 2.0
  QIRE_Start_Row             = 'Start Row';                        // since 2.0
  QIRE_Finish                = ' Finish ';                         // since 2.0
  QIRE_Finish_DataFinished   = 'Where data finished';              // since 2.0
  QIRE_Finish_Col            = 'Finish Col';                       // since 2.0
  QIRE_Finish_Row            = 'Finish Row';                       // since 2.0
  QIRE_Direction             = ' Direction ';                      // since 2.0
  QIRE_Direction_Down        = 'Down';                             // since 2.0
  QIRE_Direction_Up          = 'Up';                               // since 2.0
  QIRE_Direction_Right       = 'Right';                            // since 2.0
  QIRE_Direction_Left        = 'Left';                             // since 2.0
  QIRE_Sheet                 = ' Sheet ';                          // since 2.0
  QIRE_Sheet_Default         = 'Default Sheet';                    // since 2.0
  QIRE_Sheet_Custom          = 'Custom Sheet';                     // since 2.0
  QIRE_Sheet_Custom_Number   = 'Sheet Number';                     // since 2.0
  QIRE_Sheet_Custom_Name     = 'Sheet Name';                       // since 2.0

  QIW_Access_Table           = 'I would like to import data from ' +
                               'a table';                          // since 1.4
  QIW_Access_SQL             = 'I would like to import data from ' +
                               'an SQL query';                     // since 1.4
  QIW_Access_SQL_Load        = 'Load query from file';             // since 1.4
  QIW_Access_SQL_Save        = 'Save query to file';               // since 1.4
  QIW_Access_Add             = 'Add';                              // since 1.4
  QIW_Access_Remove          = 'Remove';                           // since 1.4
  QIW_Access_AutoFill        = 'Auto fill';                        // since 1.4
  QIW_Access_Clear           = 'Clear';                            // since 1.4
  QIW_Access_CustomQuery     = 'Custom query';
  QIW_Access_Tip             = 'Click the "Add" button to set '    +
                               'the accordance between the '       +
                               'imported column and the table '    +
                               'field or click the "Remove" '      +
                               'button to remove one.';            // since 1.4

implementation

end.
