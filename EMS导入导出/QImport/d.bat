echo off

rem ---------------
rem Local variables
rem ---------------

set D_VER=%1

set DCC=%2
set BPL=%3
set DCP=%4

set CURDIR=%CD%


set inst_log=%5

set CPPFILES=%6

if "%1"=="D19" set ADDCC="-JL"
if "%1"=="D18" set ADDCC="-JL"
if "%1"=="D17" set ADDCC="-JL"
if "%1"=="D16" set ADDCC="-JL"
if "%1"=="D15" set ADDCC="-JL"
if "%1"=="D14" set ADDCC="-JL"
if "%1"=="D12" set ADDCC="-JL"
if "%1"=="D11" set ADDCC="-JL"


if not defined DCC goto ParamsError
if not defined BPL goto ParamsError
if not defined DCP goto ParamsError


if not exist %BPL% mkdir %BPL%
if not exist %DCP% mkdir %DCP%
if not exist %CPPFILES% mkdir %CPPFILES%

if defined CPPFILES set CPPFILES1="-NO"%CPPFILES%
if defined CPPFILES set CPPFILES2="-NB"%CPPFILES%
if defined CPPFILES set CPPFILES3="-NH"%CPPFILES%


if exist %BPL%\QImport3??_%D_VER%.bpl del %BPL%\QImport3??_%D_VER%.bpl /Q
if exist %DCP%\QImport3??_%D_VER%.dcp del %DCP%\QImport3??_%D_VER%.dcp /Q

rem ------------------
rem Deleting old files
rem ------------------
if exist error_%D_VER%.log del error_%D_VER%.log

echo ------------------------------- >> error_%D_VER%.log
echo  %D_VER% >> error_%D_VER%.log
echo ------------------------------- >> error_%D_VER%.log


if exist *.~?? del *.~?? > nul
if exist *.dcu del *.dcu > nul

rem --------------------------
rem Compiling run-time package
rem --------------------------

%DCC% %ADDCC% QImport3RT_%D_VER%.dpk -B %CPPFILES1% %CPPFILES2% %CPPFILES3% >> error_%D_VER%.log
if errorlevel=1 goto RTCommonError
echo QImport3RT_%D_VER%.dpk was compiled successfully! >> error_%D_VER%.log

rem -----------------------------
rem Compiling design-time package
rem -----------------------------

%DCC% %ADDCC% QImport3DT_%D_VER%.dpk -B %CPPFILES1% %CPPFILES2% %CPPFILES3% >> error_%D_VER%.log
if errorlevel=1 goto DTCommonError          
echo QImport3DT_%D_VER%.dpk was compiled successfully! >> error_%D_VER%.log

rem --------------------
rem Copying bpl files
rem --------------------

copy *%D_VER%.bpl %BPL% /Y > nul
copy *%D_VER%.dcp %DCP% /Y > nul

rem -----------------------
rem Deleting needless files
rem -----------------------

if exist *.bpl del *.bpl > nul
if exist *.dcp del *.dcp > nul
if exist *.dcu del *.dcu > nul


goto End

:RTCommonError
echo QImport3RT_%D_VER%.dpk was compiled with errors! >> error_%D_VER%.log
echo QImport3RT_%D_VER%.dpk >>%inst_log%
rem "error_%D_VER%.log"
exit 10

:DTCommonError
echo QImport3DT_%D_VER%.dpk was compiled with errors! >> error_%D_VER%.log
echo QImport3DT_%D_VER%.dpk >>%inst_log%
rem "error_%D_VER%.log"
exit 20

:ParamsError
echo %D_VER% >> error_%D_VER%.log
echo Empty params! >> error_%D_VER%.log
exit 30

:End
