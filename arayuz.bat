@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d %~dp0
set PY=python
where py >nul 2>nul
if not errorlevel 1 set PY=py
%PY% uevcb_rapor.py
if errorlevel 1 pause
