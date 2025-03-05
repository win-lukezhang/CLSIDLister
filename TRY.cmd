@echo off
cls
title CLSIDLister
:MAIN
echo Input CLSID:
echo Example: {00000000-0000-0000-0000-000000000000}
set /p "CLSID=>"
explorer.exe shell:::%CLSID%
goto MAIN
echo You should not see this
pause
exit