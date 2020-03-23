echo on

rem ---------------
rem Local variables
rem ---------------

set OBJ_DIR=%1

set BCB_VER=%2

set RsVarsBat=%3

set BPL=%4
set LIB=%5

set inst_log=%6

rem ------------------
rem Deleting old files
rem ------------------

if exist error_%BCB_VER%.log del error_%BCB_VER%.log

@if not exist %BPL% mkdir %BPL%
@if not exist %LIB% mkdir %LIB%

@if exist %LIB%\QImport3??_%BCB_VER%.lib del %LIB%\QImport3??_%BCB_VER%.lib /Q
@if exist %LIB%\QImport3??_%BCB_VER%.bpi del %LIB%\QImport3??_%BCB_VER%.bpi /Q
@if exist %BPL%\QImport3??_%BCB_VER%.bpl del %BPL%\QImport3??_%BCB_VER%.bpl /Q

if exist *.#?? del *.#?? > nul
if exist *.~?? del *.~?? > nul
if exist *.bpl del *.bpl > nul
if exist *.dcp del *.dcp > nul
if exist *.dcu del *.dcu > nul
if exist *.obj del *.obj > nul
if exist *.hpp del *.hpp > nul
if exist *.lib del *.lib > nul
if exist *.bpi del *.bpi > nul
if exist *.tds del *.tds > nul
if exist *.map del *.map > nul
if exist *.il? del *.il? > nul

rem ------------------
rem Set Framework Path
rem ------------------

call %RsVarsBat%

rem --------------------------
rem Compiling run-time package
rem --------------------------

if exist QImport3RT_%BCB_VER%.cbproj MSBUILD.EXE /p:Config=Setup QImport3RT_%BCB_VER%.cbproj > error_%BCB_VER%.log
if errorlevel=1 goto RTCommonError
echo QImport3RT_%BCB_VER%.cbproj was builded successfully! >> error_%BCB_VER%.log

rem -----------------------------
rem Compiling design-time package
rem -----------------------------

if exist QImport3DT_%BCB_VER%.cbproj MSBUILD.EXE /p:Config=Setup QImport3DT_%BCB_VER%.cbproj > error_%BCB_VER%.log
if errorlevel=1 goto DTCommonError
echo QImport3DT_%BCB_VER%.cbproj was builded successfully! >> error_%BCB_VER%.log

rem --------------------
rem Copying result files
rem --------------------

copy *%BCB_VER%.bpl %BPL% /Y > nul
copy *%BCB_VER%.bpi %LIB% /Y > nul
copy *%BCB_VER%.lib %LIB% /Y > nul

rem --------------------
rem Save compiled files
rem --------------------

if not exist ..\%OBJ_DIR% mkdir ..\%OBJ_DIR%

if exist *.bpi copy *.bpi ..\%OBJ_DIR% > nul
if exist *.lib copy *.lib ..\%OBJ_DIR% > nul
if exist *.dcu copy *.dcu ..\%OBJ_DIR% > nul
if exist *.obj copy *.obj ..\%OBJ_DIR% > nul
if exist *.hpp copy *.hpp ..\%OBJ_DIR% > nul
if exist *.dfm copy *.dfm ..\%OBJ_DIR% > nul
if exist *.r?? copy *.r?? ..\%OBJ_DIR% > nul

rem -----------------------
rem Deleting needless files
rem -----------------------

if exist *.mak del *.mak > nul
if exist *.#?? del *.#?? > nul
if exist *.~?? del *.~?? > nul
if exist *.bpl del *.bpl > nul
if exist *.dcp del *.dcp > nul
if exist *.dcu del *.dcu > nul
if exist *.obj del *.obj > nul
if exist *.hpp del *.hpp > nul
if exist *.lib del *.lib > nul
if exist *.bpi del *.bpi > nul
if exist *.tds del *.tds > nul
if exist *.map del *.map > nul
if exist *.il? del *.il? > nul

goto End

:RTCommonError
echo QImport3RT_%BCB_VER%.cbproj was compiled with errors! >> error_%BCB_VER%.log
echo QImport3RT_%BCB_VER%.cbproj >>%inst_log%
rem "error_%BCB_VER%.log"
exit 10

:DTCommonError
echo QImport3DT_%BCB_VER%.cbproj was compiled with errors! >> error_%BCB_VER%.log
echo QImport3DT_%BCB_VER%.cbproj >>%inst_log%
rem "error_%BCB_VER%.log"
exit 20

:End
