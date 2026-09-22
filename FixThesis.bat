@echo off
chcp 65001 >nul
title FixThesis
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0FixThesis.ps1"
pause