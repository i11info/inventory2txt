@echo off
REM !run_collect_info.bat
REM This batch file will start the PowerShell script with elevated privileges

powershell -Command "Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File \"%~dp0!_collect_info.ps1\"' -Verb RunAs"
