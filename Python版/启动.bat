@echo off
cd /d "%~dp0"
if exist "日记.exe" (
  start "" "日记.exe"
) else (
  start "" ".venv\Scripts\pythonw.exe" "日记.pyw"
)
