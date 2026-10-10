@echo off
title LabSnapWord - Build and Run
echo Building and launching LabSnapWord...
dotnet run --project "%~dp0LabSnapWord\LabSnapWord.csproj" -c Release
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Could not build/run. Make sure .NET 8 SDK is installed: https://dotnet.microsoft.com/download/dotnet/8.0
    pause
)
