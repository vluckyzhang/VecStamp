@echo off
rem VecStamp - Windows install / uninstall (double-click to run)
title VecStamp Setup
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0installer\windows\VecStamp-Setup.ps1" %*
if errorlevel 1 pause
