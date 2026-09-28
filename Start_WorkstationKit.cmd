@echo off
title Workstation Kit

set "PS1=%~dp0WorkstationKit.ps1"

if not exist "%PS1%" (
    echo.
    echo WorkstationKit.ps1 is not in this folder.
    echo Keep both files together:
    echo.
    echo    Start_WorkstationKit.cmd   ^(this launcher^)
    echo    WorkstationKit.ps1         ^(the application^)
    echo.
    pause
    exit /b 1
)

start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%PS1%"
exit /b
