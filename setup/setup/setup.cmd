@echo off
cd /d %~dp0
@powershell -command "Start-Process -Verb runas addinScript.bat"
