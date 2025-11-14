@echo off
echo Running Cash Sheet Autofiller...
echo.

REM This runs your main python script
python main.py

echo.
echo Process finished.
echo Opening casheet folder for review...

REM This opens your casheet folder.
REM Note: The first "" is a dummy title to handle paths with spaces.
start "" "C:\Users\admin\OneDrive\Nam career\Chartwells\Auto-fill casheet\casheet"

echo Press any key to exit this window.
pause > nul