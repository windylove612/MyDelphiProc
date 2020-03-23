//---------------------------------------------------------------------------

#include <basepch.h>
#pragma hdrstop
USEFORMNS("QImport3Wizard.pas", Qimport3wizard, QImport3WizardF);
USEFORMNS("fuQImport3ReplacementEdit.pas", Fuqimport3replacementedit, fmQImport3ReplacementEdit);
USEFORMNS("fuQImport3ProgressDlg.pas", Fuqimport3progressdlg, fmQImport3ProgressDlg);
USEFORMNS("fuQImport3XLSRangeEdit.pas", Fuqimport3xlsrangeedit, fmQImport3XLSRangeEdit);
USEFORMNS("fuQImport3Loading.pas", Fuqimport3loading, fmQImport3Loading);
USEFORMNS("fuQImport3License.pas", Fuqimport3license, fmQImport3License);
USEFORMNS("fuQImport3About.pas", Fuqimport3about, fmQImport3About);
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
