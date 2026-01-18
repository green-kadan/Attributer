@echo off
cd /d %~dp0

if "%PROCESSOR_ARCHITECTURE%" == "x86" (
start "" msiexec /package x86\Attributer.msi
goto :EOF
)

if "%PROCESSOR_ARCHITECTURE%" == "AMD64" (
start "" msiexec /package x64\Attributer.msi
goto :EOF
)

:EOF
echo on
