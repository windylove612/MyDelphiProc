unit QImport3XLSConsts;

interface

const
  sFileNameNotDefined     = 'File name is not defined';
  sFileDoesNotExist       = 'File %s does not exist';
  sFileIsNotExcelWorkbook = 'File %s is not an Excel Workbook';

  sErrorReadingRecord     = 'Error reading record';
  sExcelInvalid           = 'Error reading Excel records. File invalid';
  sInvalidContinue        = 'Trying to add more than one CONTINUE record to the same record';
  sInvalidVersion         = 'File is saved with an invalid Excel version';
  sWrongExcelRecord       = 'Record with ID 0x%x is invalid';
  sInvalidStringRecord    = 'Invalid String record';

  sInvalidRow             = 'Invalid Row (%d)';
  sInvalidCol             = 'Invalid Column (%d)';

  sInvalidCellValue       = 'Invalid cell value: ''%s''';
  sInvalidErrStr          = 'Invalid cell error code: ''%s''';
  sShrFmlaNotFound        = 'Shared Formula not found';
  sBadToken               = 'Error reading formula, Token Id: 0x%2x is not supported';
  sBadFormula             = 'Formula at Row %d, column %d is not supported. Token Id: 0x%2x';

  sCellAccessError        = 'Cannot access cell ''%s'' as type %s';
  sValueMissing           = 'Value is missing';
  sEmptyStack             = 'EmptyStack';
  sInvalidAreaArgument    = 'Invalid Reference(Area) argument';
  sInvalidArraySize       = 'Invalid variant array size';
  sUnknownPTG             = 'Unknown ptg[%.2X]';

implementation

end.
