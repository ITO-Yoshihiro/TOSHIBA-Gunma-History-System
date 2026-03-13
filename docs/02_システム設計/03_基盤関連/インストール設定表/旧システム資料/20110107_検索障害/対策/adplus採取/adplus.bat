@echo off

set OUT_PATH="C:\WORK"
set BIN_PATH="C:\Program Files\Debugging Tools for Windows (x86)"

cd %BIN_PATH%

cscript adplus.vbs -hang -iis -o %OUT_PATH% -quiet

exit