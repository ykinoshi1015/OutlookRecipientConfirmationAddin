@echo off
cd /d %~dp0

rem -----�Z�L�����e�B�ؖ���-----
certutil -addstore ROOT OutlookRecipientConfirmationAddin.cer

rem -----Microsoft Visual C++ 2010 �ĔЕz�\�p�b�P�[�W (x86)-----
vcredist_x86.exe /Setup /passive /promptrestart

call addinInstaller.exe

rem -----dll���Ăяo��powershell�����s-----
powershell -NoProfile -ExecutionPolicy Unrestricted .\sample.ps1
