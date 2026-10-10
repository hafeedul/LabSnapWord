@echo off
title LabSnap Word - Build Standalone App & Desktop Shortcut
echo Building standalone Windows app (LabSnapWord.exe with custom icon)...
dotnet publish "%~dp0LabSnapWord\LabSnapWord.csproj" -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:EnableCompressionInSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o "%~dp0dist"
if %ERRORLEVEL% EQU 0 (
    copy /Y "%~dp0dist\LabSnapWord.exe" "%~dp0dist\LabSnapWord-Setup.exe" >nul
    echo.
    echo [SUCCESS] Standalone app & installer created in: %~dp0dist\
) else (
    echo.
    echo [ERROR] Build failed. Ensure the .NET 8 SDK is installed.
)
pause
