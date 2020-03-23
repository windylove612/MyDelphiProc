//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
USEPACKAGE("vcl50.bpi");
USEPACKAGE("Vcldb50.bpi");
USEPACKAGE("Vclx50.bpi");
USEFORMNS("fuQImport3ProgressDlg.pas", Fuqimport3progressdlg, fmQImport3ProgressDlg);
USEUNIT("QImport3XLS.pas");
USEUNIT("QImport3.pas");
USEUNIT("QImport3ASCII.pas");
USEUNIT("QImport3DBF.pas");
USEUNIT("QImport3XLSMapParser.pas");
USEUNIT("QImport3XLSUtils.pas");
USEUNIT("QImport3XLSCalculate.pas");
USEUNIT("QImport3XLSCommon.pas");
USEUNIT("QImport3XLSConsts.pas");
USEUNIT("QImport3XLSFile.pas");
USEFORMNS("fuQImport3About.pas", Fuqimport3about, fmQImport3About);
USEFORMNS("fuQImport3Loading.pas", Fuqimport3loading, fmQImport3Loading);
USEUNIT("QImport3Common.pas");
USEUNIT("QImport3DataSet.pas");
USEFORMNS("fuQImport3ReplacementEdit.pas", Fuqimport3replacementedit, fmQImport3ReplacementEdit);
USEFORMNS("fuQImport3XLSRangeEdit.pas", Fuqimport3xlsrangeedit, fmQImport3XLSRangeEdit);
USEUNIT("QImport3InfoPanel.pas");
USEUNIT("QImport3TXTView.pas");
USEUNIT("QImport3XML.pas");
USEUNIT("QImport3DBFFile.pas");
USEFORMNS("QImport3Wizard.pas", Qimport3wizard, QImport3WizardF);
USEFORMNS("fuQImport3License.pas", Fuqimport3license, fmQImport3License);
USEUNIT("QImport3WideStrUtils.pas");
USEUNIT("QImport3CustomControl.pas");
USEUNIT("QImport3WideStringCanvas.pas");
USEUNIT("QImport3WideStringGrid.pas");
USEUNIT("QImport3WideStrings.pas");
USEUNIT("QImport3EZDSLTHD.PAS");
USEUNIT("QImport3EZDSLBSE.PAS");
USEUNIT("QImport3EZDSLCTS.PAS");
USEUNIT("QImport3EZDSLHSH.PAS");
USEUNIT("QImport3EZDSLLST.PAS");
USEUNIT("QImport3EZDSLSUP.PAS");
USEUNIT("QImport3zutil.pas");
USEUNIT("QImport3ADLER.PAS");
USEUNIT("QImport3crc.pas");
USEUNIT("QImport3gzio.pas");
USEUNIT("QImport3infblock.pas");
USEUNIT("QImport3infcodes.pas");
USEUNIT("QImport3inffast.pas");
USEUNIT("QImport3inftrees.pas");
USEUNIT("QImport3infutil.pas");
USEUNIT("QImport3libdatei.pas");
USEUNIT("QImport3trees.pas");
USEUNIT("QImport3uCommon.pas");
USEUNIT("QImport3Unzip.pas");
USEUNIT("QImport3zcompres.pas");
USEUNIT("QImport3ZDEFLATE.PAS");
USEUNIT("QImport3ZINFLATE.PAS");
USEUNIT("QImport3zip.pas");
USEUNIT("QImport3ZipMcpt.pas");
USEUNIT("QImport3ziputils.pas");
USEUNIT("QImport3Zlib110.pas");
USEUNIT("QImport3zuncompr.pas");
USEUNIT("QImport3BaseDocumentFile.pas");
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
