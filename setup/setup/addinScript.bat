@echo off
cd /d %~dp0

rem -----DLL��PowerShell�̂̃u���b�N����-----
start /wait streams.exe -d "DoNotDisableAddinUpdater.dll.deploy"
start /wait streams.exe -d "callRegistryDll.ps1"

rem -----�Z�L�����e�B�ؖ���-----
certutil -addstore ROOT OutlookRecipientConfirmationAddin.cer

rem -----Microsoft Visual C++ 2010 SP1 �ĔЕz�\�p�b�P�[�W (x86)-----
vcredist_x86.exe /Setup /passive /promptrestart

rem -----�C���X�g�[���[�Ăяo��-----
call addinInstaller.exe

rem -----dll���Ăяo��powershell�����s-----
powershell -NoProfile -ExecutionPolicy Unrestricted .\callRegistryDll.ps1