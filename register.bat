@echo off
setlocal

:: Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"
set "DLL_PATH=%SCRIPT_DIR%build\bin\Release\NewFolderFromFiles.dll"

if "%1"=="unregister" goto unregister

echo Registering NewFolderFromFiles.dll...
echo DLL: %DLL_PATH%
%SystemRoot%\System32\regsvr32.exe "%DLL_PATH%"
goto end

:unregister
echo Unregistering NewFolderFromFiles.dll...
%SystemRoot%\System32\regsvr32.exe /u "%DLL_PATH%"

:end
endlocal

pause