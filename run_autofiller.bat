@echo off
echo Running Cash Sheet Autofiller...
echo.

REM Run the new .exe file instead of the python script
main.exe

echo.
echo Process finished.
echo Opening casheet folder for review...
start "" "casheet"

echo Press any key to exit this window.
pause > nul