@echo off

rem ---------------
rem Local variables
rem ---------------

@set OBJ_DIR=%1

@set BCB_VER=%2

@set MAKE=%3
@set bpr2mak=%4
@set BPL=%5
@set LIB=%6

@set inst_log=%7

@if not defined MAKE goto ParamsError
@if not defined bpr2mak goto ParamsError
@if not defined BPL goto ParamsError
@if not defined LIB goto ParamsError

@if not exist %BPL% mkdir %BPL%
@if not exist %LIB% mkdir %LIB%

@if exist %LIB%\QImport3??_%BCB_VER%.lib del %LIB%\QImport3??_%BCB_VER%.lib /Q
@if exist %LIB%\QImport3??_%BCB_VER%.bpi del %LIB%\QImport3??_%BCB_VER%.bpi /Q
@if exist %BPL%\QImport3??_%BCB_VER%.bpl del %BPL%\QImport3??_%BCB_VER%.bpl /Q

rem ------------------
rem Deleting old files
rem ------------------

@if exist error_%BCB_VER%.log del error_%BCB_VER%.log

echo ------------------------------- >> error_%BCB_VER%.log
echo  %BCB_VER% >> error_%BCB_VER%.log
echo ------------------------------- >> error_%BCB_VER%.log

@if exist *.~?? del *.~?? > nul
@if exist *.bpl del *.bpl > nul
@if exist *.dcp del *.dcp > nul
@if exist *.dcu del *.dcu > nul
@if exist *.obj del *.obj > nul
@if exist *.hpp del *.hpp > nul
@if exist *.lib del *.lib > nul
@if exist *.bpi del *.bpi > nul
@if exist *.tds del *.tds > nul

rem -------------------
rem Creating make files
rem -------------------

rem bpr2mak -q -t%BCB_VER%.bmk QImport3RT_%BCB_VER%.bpk > nul
rem bpr2mak -q -t%BCB_VER%.bmk QImport3DT_%BCB_VER%.bpk > nul
@%bpr2mak% QImport3RT_%BCB_VER%.bpk > nul
@%bpr2mak% QImport3DT_%BCB_VER%.bpk > nul

rem --------------------------
rem Compiling run-time package
rem --------------------------

@if exist QImport3RT_%BCB_VER%.mak %MAKE% -fQImport3RT_%BCB_VER%.mak -B >> error_%BCB_VER%.log
@if errorlevel=1 goto RTCommonError
@echo QImport3RT_%BCB_VER%.bpk was compiled successfully! >> error_%BCB_VER%.log

rem -----------------------------
rem Compiling design-time package
rem -----------------------------

@if exist QImport3DT_%BCB_VER%.mak %MAKE% -fQImport3DT_%BCB_VER%.mak -B >> error_%BCB_VER%.log
@if errorlevel=1 goto DTCommonError
@echo QImport3DT_%BCB_VER%.bpk was compiled successfully! >> error_%BCB_VER%.log

rem --------------------
rem Copying result files
rem --------------------

@copy *%BCB_VER%.bpl %BPL% /Y > nul
@copy *%BCB_VER%.bpi %LIB% /Y > nul
@copy *%BCB_VER%.lib %LIB% /Y > nul

rem --------------------
rem Save compiled files
rem --------------------

@if not exist ..\%OBJ_DIR% mkdir ..\%OBJ_DIR%

@if exist *.bpl copy *.bpl ..\%OBJ_DIR% > nul
@if exist *.bpi copy *.bpi ..\%OBJ_DIR% > nul
@if exist *.lib copy *.lib ..\%OBJ_DIR% > nul
@if exist *.dcu copy *.dcu ..\%OBJ_DIR% > nul
@if exist *.obj copy *.obj ..\%OBJ_DIR% > nul
@if exist *.hpp copy *.hpp ..\%OBJ_DIR% > nul
@if exist *.dfm copy *.dfm ..\%OBJ_DIR% > nul
@if exist *.res copy *.res ..\%OBJ_DIR% > nul
@if exist *.r32 copy *.r32 ..\%OBJ_DIR% > nul

rem -----------------------
rem Deleting needless files
rem -----------------------

@if exist *.mak del *.mak > nul
@if exist *.dcu del *.dcu > nul
@if exist *.obj del *.obj > nul
@if exist *.hpp del *.hpp > nul
@if exist *.lib del *.lib > nul
@if exist *.bpi del *.bpi > nul
@if exist *.bpl del *.bpl > nul
@if exist *.tds del *.tds > nul

goto End

:RTCommonError
@echo QImport3RT_%BCB_VER%.bpk was compiled with errors! >> error_%BCB_VER%.log
@echo QImport3RT_%BCB_VER%.bpk >>%inst_log%
rem @"error_%BCB_VER%.log"
@exit 10

:DTCommonError
@echo QImport3DT_%BCB_VER%.bpk was compiled with errors! >> error_%BCB_VER%.log
@echo QImport3DT_%BCB_VER%.bpk >>%inst_log%
rem @"error_%BCB_VER%.log"
@exit 20

:ParamsError
@echo %BCB_VER% >> error_%BCB_VER%.log
@echo Empty params! >> error_%BCB_VER%.log
@exit 30

:End