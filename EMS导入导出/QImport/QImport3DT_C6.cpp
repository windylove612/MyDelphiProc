//---------------------------------------------------------------------------

#include <basepch.h>
#pragma hdrstop
USEFORMNS("fuQImport3XMLEditor.pas", FuqImport3xmleditor, fmQImport3XMLEditor);
USEFORMNS("fuQImport3CSVEditor.pas", FuqImport3csveditor, fmQImport3CSVEditor);
USEFORMNS("fuQImport3DataSetEditor.pas", FuqImport3dataseteditor, fmQImport3DataSetEditor);
USEFORMNS("fuQImport3DBFEditor.pas", FuqImport3dbfeditor, fmQImport3DBFEditor);
USEFORMNS("fuQImport3FormatsEditor.pas", FuqImport3formatseditor, fmQImport3FormatsEditor);
USEFORMNS("fuQImport3TXTEditor.pas", FuqImport3txteditor, fmQImport3TXTEditor);
USEFORMNS("fuQImport3XLSEditor.pas", FuqImport3xlseditor, fmQImport3XLSEditor);
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
