@echo off

py -m pip install --upgrade pip
py -m pip install -r %~dp0/requirements.txt

echo.
if %ERRORLEVEL% EQU 0 (
    echo Installation completed successfully.
    echo You're good to go! Run 'run.pyw' to start the program.
) else (
    echo Installation failed!
    echo Please make sure to read all the info in the README before running this script.
    echo Otherwise, email the error message above to the current maintainer. 
)

echo.
echo Press any key to continue...
pause > nul