//---------------------------------------------------------------------------

#include <basepch.h>
#pragma hdrstop
USEFORMNS("fuQImport3CSVEditor.pas", Fuqimport3csveditor, fmQImport3CSVEditor);
USEFORMNS("fuQImport3DataSetEditor.pas", Fuqimport3dataseteditor, fmQImport3DataSetEditor);
USEFORMNS("fuQImport3DBFEditor.pas", Fuqimport3dbfeditor, fmQImport3DBFEditor);
USEFORMNS("fuQImport3FormatsEditor.pas", Fuqimport3formatseditor, fmQImport3FormatsEditor);
USEFORMNS("fuQImport3TXTEditor.pas", Fuqimport3txteditor, fmQImport3TXTEditor);
USEFORMNS("fuQImport3XLSEditor.pas", Fuqimport3xlseditor, fmQImport3XLSEditor);
USEFORMNS("fuQImport3HTMLEditor.pas", Fuqimport3htmleditor, fmQImport3HTMLEditor);
USEFORMNS("fuQImport3XMLDocEditor.pas", Fuqimport3xmldoceditor, fmQImport3XMLDocEditor);
USEFORMNS("fuQImport3XMLEditor.pas", Fuqimport3xmleditor, fmQImport3XMLEditor);
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
