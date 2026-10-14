@echo off
cd /d "%~dp0"
python calculator_app.py
if errorlevel 1 pause
