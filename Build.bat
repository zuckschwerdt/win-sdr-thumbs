@echo off

:: Check if running as administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo This script must be run as Administrator.
    echo Please right-click on the batch file and select "Run as administrator".
    echo.
    pause
    exit /b 1
)

:: Unregister the previous DLL version
regsvr32 /u "%~dp0target\x86_64-pc-windows-msvc\release\win_svg_thumbs.dll"

:: (Requires IObit Unlocker to be installed) We can't delete the DLL if it's in use, which is often the case even after unregistering it. So this will unlock it and delete it.
"C:\Program Files (x86)\IObit\IObit Unlocker\IObitUnlocker.exe" /Delete /Normal "%~dp0target\x86_64-pc-windows-msvc\release\win_svg_thumbs.dll"

:: Run the build. Tried to to make it run as a non-admin, but couldn't get it to work.
cd /d %~dp0
cargo build --release --target=x86_64-pc-windows-msvc

:: Re-register the new DLL version
regsvr32 "%~dp0target\x86_64-pc-windows-msvc\release\win_svg_thumbs.dll"
