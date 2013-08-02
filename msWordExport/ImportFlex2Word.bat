@echo off
REM ImportFlex2Word.bat - 2-Aug-2013 Greg Trihus
set myProg=\SIL\msWordExport\msWordExport.exe
set progDir=C:\Program Files
if exist "%progDir%%myProg%" goto foundIt
set progDir=%ProgramFiles(x86)%
if exist "%progDir%%myProg%" goto foundIt
set progDir=%ProgramFiles%
if exist "%progDir%%myProg%" goto fountIt
echo msWordExport.exe not found
goto done

:foundIt
@echo on
"%progDir%%myProg%" -v -o=main.doc -t="My Title" "%1"
@echo off
:done
pause