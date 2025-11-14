@echo off
echo Running Cash Sheet Autofiller...
echo.

REM --- FIX: Tell Python to run the script inside src ---
python src/main.py

echo.
echo Process finished.
echo Opening casheet folder for review...

REM This still works perfectly because 'casheet' is in the same folder as the .bat file
start "" "casheet"

echo Press any key to exit this window.
pause > nul