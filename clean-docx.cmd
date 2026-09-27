@echo off
chcp 65001 >nul
if "%~1"=="" (
  echo Drag one or more .docx files onto this file to clean them.
  echo Usage: clean-docx.cmd "C:\path\file.docx"
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0docx-clean.ps1" %*
echo.
pause
