//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
USEPACKAGE("vcl50.bpi");
USEPACKAGE("Vcldb50.bpi");
USEFORMNS("fuQImport3XMLEditor.pas", Fuqimport3xmleditor, fmQImport3XMLEditor);
USEUNIT("QImport3Reg.pas");
USEFORMNS("fuQImport3CSVEditor.pas", Fuqimport3csveditor, fmQImport3CSVEditor);
USEFORMNS("fuQImport3DataSetEditor.pas", Fuqimport3dataseteditor, fmQImport3DataSetEditor);
USEFORMNS("fuQImport3DBFEditor.pas", Fuqimport3dbfeditor, fmQImport3DBFEditor);
USEFORMNS("fuQImport3FormatsEditor.pas", Fuqimport3formatseditor, fmQImport3FormatsEditor);
USEFORMNS("fuQImport3TXTEditor.pas", Fuqimport3txteditor, fmQImport3TXTEditor);
USEFORMNS("fuQImport3XLSEditor.pas", Fuqimport3xlseditor, fmQImport3XLSEditor);
USEPACKAGE("QImport3RT_C5.bpi");
USEPACKAGE("Vclx50.bpi");
//---------------------------------------------------------------------------
#pragma package(smart_init)
//---------------------------------------------------------------------------

//   Package source.
//---------------------------------------------------------------------------

#pragma argsused
int WINAPI DllEntryPoint(HINSTANCE hinst, unsigned long reason, void*)
{
        return 1;
}
//---------------------------------------------------------------------------
