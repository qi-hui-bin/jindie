@echo off
setlocal
chcp 65001 >nul
title Jindie Demo
if not exist "%~dp0runtime\python.exe" (
    echo Jindie runtime is missing. Extract the entire ZIP before starting.
    pause
    exit /b 2
)
"%~dp0runtime\python.exe" -m jindie.launcher --bundle-root "%~dp0." --demo
if errorlevel 1 (
    pause
    exit /b 2
)
