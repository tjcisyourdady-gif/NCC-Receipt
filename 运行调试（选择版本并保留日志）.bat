@echo off
cd /d "%~dp0"
setlocal EnableExtensions EnableDelayedExpansion
set "COUNT=0"
for %%F in ("*.exe") do (
    if exist "%%~fF" (
        set /a COUNT+=1
        set "APP_!COUNT!=%%~fF"
        echo !COUNT!. %%~nxF
    )
)
if "%COUNT%"=="0" (
    echo No EXE was found in this folder.
    pause
    exit /b 2
)
echo.
set /p "PICK=Select the program number to debug: "
call set "APP=%%APP_%PICK%%%"
if not defined APP (
    echo Invalid selection.
    pause
    exit /b 3
)
echo.
echo Starting: %APP%
"%APP%"
set "APP_EXIT_CODE=%ERRORLEVEL%"
echo.
echo Program exited with code: %APP_EXIT_CODE%
pause
exit /b %APP_EXIT_CODE%
