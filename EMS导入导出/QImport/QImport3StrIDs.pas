unit QImport3StrIDs;

interface

const
  QIRes = 40000;

  // Errors
  QIE_NoDataSet = QIRes + 2000;

  QIE_UnknownImportSource = QIRes + 2001;
  QIE_WrongFileType = QIRes + 2002;
  QIE_UnknownFieldType = QIRes + 2003;
  QIE_WrongDateFormat = QIRes + 2004;
  QIE_ListIndexOutOfBounds = QIRes + 2005;
  QIE_GridCaptionPos = QIRes + 2006;
  QIE_IsNotBoolean = QIRes + 4224;

  QIE_NoDBGrid = QIRes + 2010;
  QIE_NoListView = QIRes + 2020;
  QIE_NoStringGrid = QIRes + 2030;
  QIE_NoSource = QIRes + 2040;
  QIE_NoFileName = QIRes + 2050;
  QIE_FileNotExists = QIRes + 2060;
  QIE_MappingEmpty = QIRes + 2070;
  QIE_CanceledByUser = QIRes + 2080;
  QIE_MapMissing = QIRes + 2090;
  QIE_FieldNotFound = QIRes + 2100;
  QIE_SourceFieldNotFound = QIRes + 2110;
  QIE_DataSetClosed = QIRes + 2120;
  QIE_AllowedImportsEmpty = QIRes + 2130;
  QIE_KeyColumnsNotDefined = QIRes + 2140;
  QIE_KeyColumnNotFound = QIRes + 2150;
  QIE_XMLFieldListEmpty = QIRes + 2160;
  QIE_NeedDefineTextToFind = QIRes + 2170;
  QIE_XlsxMapError = QIRes + 4223;

  // Filters
  QIF_TXT = QIRes + 2180;
  QIF_ASCII = QIRes + 2190;
  QIF_CSV = QIRes + 2200;
  QIF_DBF = QIRes + 2210;
  QIF_XML = QIRes + 2220;
  QIF_ODT = QIRes + 2225;
  QIF_XLS = QIRes + 2230;
  QIF_XLSX = QIRes + 2231;
  QIF_DOCX = QIRes + 2232;
  QIF_ODS = QIRes + 2235;
  QIF_Access = QIRes + 2240;
  QIF_HTML = QIRes + 2245;
  QIF_ALL = QIRes + 2250;
  QIF_TEMPLATE = QIRes + 2260;

  // Loading
  QIL_Loading = QIRes + 2270;

  // Defaults
  QID_BooleanTrue = QIRes + 2280;
  QID_BooleanFalse = QIRes + 2290;
  QID_NullValue = QIRes + 2300;

  // Wizard
  QIW_importWizard = QIRes + 2310;
  QIW_ImportFrom = QIRes + 2320;
  QIW_ODT = QIRes + 2325;
  QIW_XLS = QIRes + 2330;
  QIW_XLSX = QIRes + 2331;
  QIW_DOCX = QIRes + 2332;
  QIW_ODS = QIRes + 2335;
  QIW_Access = QIRes + 2340;
  QIW_DBF = QIRes + 2350;
  QIW_HTML = QIRes + 2355;
  QIW_XML = QIRes + 2360;
  QIW_XMLDoc = QIRes + 2361;
  QIW_TXT = QIRes + 2370;
  QIW_CSV = QIRes + 2380;
  QIW_Comma = QIRes + 2390;
  QIW_Quote = QIRes + 2400;
  QIW_FileName = QIRes + 2410;
  QIW_NeedCancelCaption = QIRes + 2420;
  QIW_NeedCancel = QIRes + 2430;
  QIW_Question = QIRes + 2440;
  QIW_CancelWizard = QIRes + 2450;
  QIW_TemplateFileName = QIRes + 2451;

  QIW_Help = QIRes + 2460;
  QIW_Back = QIRes + 2470;
  QIW_Next = QIRes + 2480;
  QIW_Cancel = QIRes + 2490;
  QIW_Execute = QIRes + 2500;
  QIW_Step = QIRes + 2510;
  QIW_TemplateOptions = QIRes + 2520;
  QIW_GoToLastPage = QIRes + 2530;

  QIW_AutoSaveTemplate = QIRes + 2540;

  QIW_LoadTemplate = QIRes + 2550;
  QIW_SaveTemplate = QIRes + 2560;

  QIW_CommitOptions = QIRes + 2570;
  QIW_Commit = QIRes + 2580;
  QIW_CommitAfterDone = QIRes + 2590;
  QIW_CommitRecCount = QIRes + 2600;
  QIW_Records = QIRes + 2610;
  QIW_RecordCount = QIRes + 2620;
  QIW_ImportAllRecords = QIRes + 2630;
  QIW_ImportRecCount = QIRes + 2640;
  QIW_CloseAfterImport = QIRes + 2650;
  QIW_AutoTrimValue = QIRes + 4240;
  QIW_ImportEmptyRows = QIRes + 4241;

  QIW_ImportAdvanced = QIRes + 2660;
  QIW_AddType = QIRes + 2670;
  QIW_AddType_Append = QIRes + 2680;
  QIW_AddType_Insert = QIRes + 2690;

  QIWIM_Caption = QIRes + 2700;
  QIWIM_Insert_All = QIRes + 2710;
  QIWIM_Insert_New = QIRes + 2720;
  QIWIM_Update = QIRes + 2730;
  QIWIM_Update_or_Insert = QIRes + 2740;
  QIWIM_Delete = QIRes + 2750;
  QIWIM_Delete_or_Insert = QIRes + 2760;

  QIW_AvailableColumns = QIRes + 2770;
  QIW_SelectedColumns = QIRes + 2780;

  QIW_ErrorLog = QIRes + 2790;
  QIW_EnableErrorLog = QIRes + 2800;
  QIW_ErrorLogFileName = QIRes + 2810;
  QIW_RewriteErrorLogFile = QIRes + 2820;

  QIW_ShowErrorLog = QIRes + 2830;
  QIW_ErrorLogStarted = QIRes + 2840;
  QIW_ErrorLogFinished = QIRes + 2850;
  QIW_SomeErrorsFound = QIRes + 2860;
  QIW_NoErrorsFound = QIRes + 2870;

  QIW_BaseFormats = QIRes + 2880;
  QIW_DateTimeFormats = QIRes + 2890;
  QIW_Separators = QIRes + 2900;
  QIW_DecimalSeparator = QIRes + 2910;
  QIW_ThousandSeparator = QIRes + 2920;
  QIW_ShortDateFormat = QIRes + 2930;
  QIW_LongDateFormat = QIRes + 2940;
  QIW_DateSeparator = QIRes + 2950;
  QIW_ShortTimeFormat = QIRes + 2960;
  QIW_LongTimeFormat = QIRes + 2970;
  QIW_TimeSeparator = QIRes + 2980;

  QIW_BooleanTrue = QIRes + 2990;
  QIW_BooleanFalse = QIRes + 3000;
  QIW_NullValue = QIRes + 3010;

  QIWDF_Caption = QIRes + 3020;
  QIWDF_Fields = QIRes + 3030;
  QIWDF_Tuning = QIRes + 3040;
  QIWDF_Replacements = QIRes + 3050;
  QIWDF_TextToFind = QIRes + 3060;
  QIWDF_ReplaceWith = QIRes + 3070;
  QIWDF_IgnoreCase = QIRes + 3080;
  QIWDF_IgnoreCase_Yes = QIRes + 3090;
  QIWDF_IgnoreCase_No = QIRes + 3100;
  QIWDF_AddReplacement = QIRes + 3110;
  QIWDF_EditReplacement = QIRes + 3120;
  QIWDF_DelReplacement = QIRes + 3130;
  QIWDF_GeneratorValue = QIRes + 3140;
  QIWDF_GeneratorStep = QIRes + 3150;
  QIWDF_ConstantValue = QIRes + 3160;
  QIWDF_NullValue = QIRes + 3170;
  QIWDF_DefaultValue = QIRes + 3180;
  QIWDF_LeftQuote = QIRes + 3190;
  QIWDF_RightQuote = QIRes + 3200;
  QIWDF_QuoteAction = QIRes + 3210;
  QIWDF_QuoteNone = QIRes + 3220;
  QIWDF_QuoteAdd = QIRes + 3230;
  QIWDF_QuoteRemove = QIRes + 3240;
  QIWDF_CharCase = QIRes + 3250;
  QIWDF_CharCaseNone = QIRes + 3260;
  QIWDF_CharCaseUpper = QIRes + 3270;
  QIWDF_CharCaseLower = QIRes + 3280;
  QIWDF_CharCaseUpperFirst = QIRes + 3290;
  QIWDF_CharCaseUpperWord = QIRes + 3300;
  QIWDF_CharSet = QIRes + 3310;
  QIWDF_CharSetNone = QIRes + 3320;
  QIWDF_CharSetAnsi = QIRes + 3330;
  QIWDF_CharSetOem = QIRes + 3340;
  QIWDF_JScript = QIRes + 4290;

  QIWDF_RE_Title = QIRes + 3350;
  QIWDF_RE_TextToFind = QIRes + 3360;
  QIWDF_RE_ReplaceWith = QIRes + 3370;
  QIWDF_RE_IgnoreCase = QIRes + 3380;

  QIW_TXT_Fields = QIRes + 3390;
  QIW_TXT_Fields_Pos = QIRes + 3400;
  QIW_TXT_Fields_Size = QIRes + 3410;
  QIW_TXT_Clear = QIRes + 3420;
  QIW_TXT_SkipLines = QIRes + 3430;
  QIW_TXT_Tip = QIRes + 3440;

  QIW_CSV_Fields = QIRes + 3450;
  QIW_CSV_ColNumber = QIRes + 3460;
  QIW_CSV_SkipLines = QIRes + 3470;
  QIW_CSV_AutoFill = QIRes + 3480;
  QIW_CSV_Clear = QIRes + 3490;
  QIW_CSV_GridCol = QIRes + 3500;
  QIW_CSV_Tip = QIRes + 3510;

  QIW_DBF_Add = QIRes + 3520;
  QIW_DBF_Remove = QIRes + 3530;
  QIW_DBF_AutoFill = QIRes + 3540;
  QIW_DBF_Clear = QIRes + 3550;
  QIW_DBF_SkipDeleted = QIRes + 3560;
  QIW_DBF_Tip = QIRes + 3570;

  QIW_Xlsx_Fields = QIRes + 3551;
  QIW_Xlsx_SkipRows = QIRes + 3552;
  QIW_Xlsx_AutoFill = QIRes + 3553;
  QIW_Xlsx_Clear = QIRes + 3554;
  QIW_Xlsx_GridCol = QIRes + 3555;
  QIW_Xlsx_Tip = QIRes + 3556;

  QIW_Docx_Fields = QIRes + 3561;
  QIW_Docx_SkipRows = QIRes + 3562;
  QIW_Docx_AutoFill = QIRes + 3563;
  QIW_Docx_Clear = QIRes + 3564;
  QIW_Docx_GridCol = QIRes + 3565;
  QIW_Docx_Tip = QIRes + 3566;

  QIW_HTML_Fields = QIRes + 3571;
  QIW_HTML_ColNumber = QIRes + 3572;
  QIW_HTML_TableNumber = QIRes + 3573;
  QIW_HTML_SkipLines = QIRes + 3574;
  QIW_HTML_AutoFill = QIRes + 3575;
  QIW_HTML_Clear = QIRes + 3576;
  QIW_HTML_GridCol = QIRes + 3577;
  QIW_HTML_Tip = QIRes + 3578;

  QIW_XML_Add = QIRes + 3580;
  QIW_XML_Remove = QIRes + 3590;
  QIW_XML_AutoFill = QIRes + 3600;
  QIW_XML_Clear = QIRes + 3610;
  QIW_XML_WriteOnFly = QIRes + 3620;
  QIW_XML_Tip = QIRes + 3630;

  QIW_ODS_Fields = QIRes + 3581;
  QIW_ODS_SkipRows = QIRes + 3582;
  QIW_ODS_AutoFill = QIRes + 3583;
  QIW_ODS_Clear = QIRes + 3584;
  QIW_ODS_GridCol = QIRes + 3585;
  QIW_ODS_Tip = QIRes + 3586;

  QIW_ODT_Fields = QIRes + 3591;
  QIW_ODT_SkipRows = QIRes + 3592;
  QIW_ODT_AutoFill = QIRes + 3593;
  QIW_ODT_Clear = QIRes + 3594;
  QIW_ODT_GridCol = QIRes + 3595;
  QIW_ODT_Tip = QIRes + 3596;

  QIW_XMLDoc_Fields = QIRes + 3631;
  QIW_XMLDoc_ColNumber = QIRes + 3632;
  QIW_XMLDoc_SkipLines = QIRes + 3633;
  QIW_XMLDoc_AutoFill = QIRes + 3634;
  QIW_XMLDoc_Clear = QIRes + 3635;
  QIW_XMLDoc_GridCol = QIRes + 3636;
  QIW_XMLDoc_Tip = QIRes + 3637;
  QIW_XMLDoc_XPath = QIRes + 3638;
  QIW_XMLDoc_DataLocation = QIRes + 3639;

  QIW_XLS_Fields = QIRes + 3640;
  QIW_XLS_Ranges = QIRes + 3650;
  QIW_XLS_AddRange = QIRes + 3660;
  QIW_XLS_EditRange = QIRes + 3670;
  QIW_XLS_DelRange = QIRes + 3680;
  QIW_XLS_MoveRangeUp = QIRes + 3690;
  QIW_XLS_MoveRangeDown = QIRes + 3700;
  QIW_XLS_SkipCols = QIRes + 3710;
  QIW_XLS_SkipRows = QIRes + 3720;
  QIW_XLS_AutoFillCols = QIRes + 3730;
  QIW_XLS_AutoFillRows = QIRes + 3740;
  QIW_XLS_ClearFieldRanges = QIRes + 3750;
  QIW_XLS_ClearAllRanges = QIRes + 3760;
  QIW_XLS_Tip = QIRes + 3770;

  QIRE_Caption = QIRes + 3780;
  QIRE_RangeType = QIRes + 3790;
  QIRE_RangeType_Col = QIRes + 3800;
  QIRE_RangeType_Row = QIRes + 3810;
  QIRE_RangeType_Cell = QIRes + 3820;
  QIRE_Start = QIRes + 3830;
  QIRE_Start_DataStarted = QIRes + 3840;
  QIRE_Start_Col = QIRes + 3850;
  QIRE_Start_Row = QIRes + 3860;
  QIRE_Finish = QIRes + 3870;
  QIRE_Finish_DataFinished = QIRes + 3880;
  QIRE_Finish_Col = QIRes + 3890;
  QIRE_Finish_Row = QIRes + 3900;
  QIRE_Direction = QIRes + 3910;
  QIRE_Direction_Down = QIRes + 3920;
  QIRE_Direction_Up = QIRes + 3930;
  QIRE_Direction_Right = QIRes + 3940;
  QIRE_Direction_Left = QIRes + 3950;
  QIRE_Sheet = QIRes + 3960;
  QIRE_Sheet_Default = QIRes + 3970;
  QIRE_Sheet_Custom = QIRes + 3980;
  QIRE_Sheet_Custom_Number = QIRes + 3990;
  QIRE_Sheet_Custom_Name = QIRes + 4000;

  QIW_Access_Table = QIRes + 4010;

  QIW_Access_SQL = QIRes + 4020;

  QIW_Access_SQL_Load = QIRes + 4030;
  QIW_Access_SQL_Save = QIRes + 4040;
  QIW_Access_Add = QIRes + 4050;
  QIW_Access_Remove = QIRes + 4060;
  QIW_Access_AutoFill = QIRes + 4070;
  QIW_Access_Clear = QIRes + 4080;
  QIW_Access_CustomQuery = QIRes + 4090;
  QIW_Access_Tip = QIRes + 4100;

  QIW_Access_Password = QIRes + 4130;

  QIW_ImportErrorFormat = QIRes + 4140;

  QIPD_State = QIRes + 4201;
  QIPD_Processed = QIRes + 4202;
  QIPD_Inserted = QIRes + 4203;
  QIPD_Updated = QIRes + 4204;
  QIPD_Deleted = QIRes + 4205;
  QIPD_Errors = QIRes + 4206;
  QIPD_Committed = QIRes + 4207;
  QIPD_Time = QIRes + 4208;
  QIPD_Cancel = QIRes + 4209;
  QIPD_Import_finished = QIRes + 4210;
  QIPD_OK = QIRes + 4211;
  QIPD_Importing = QIRes + 4212;
  QIPD_Preparing = QIRes + 4213;
  QIPD_Aborted = QIRes + 4214;
  QIPD_Finished = QIRes + 4215;
  QIPD_Paused = QIRes + 4216;

  QIRE_Cancel = QIRes + 4217;

  QIW_ODT_UseFirstRowAsHeader =  QIRes + 4218;

  QIW_TXT_Encoding = QIRes + 4219;

  QIW_XMLDoc_BuildTreeView = QIRes + 4220;
  QIW_XMLDoc_GetXPath = QIRes + 4221;

  QIW_DontWorkUserDefined = QIRes + 4222;

implementation

end.
