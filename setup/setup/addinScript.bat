@echo off
cd /d %~dp0

rem -----DLL��PowerShell�̂̃u���b�N����-----
start /wait streams.exe -d "DoNotDisableAddinUpdater.dll.deploy"
start /wait streams.exe -d "callRegistryDll.ps1"

rem -----�Z�L�����e�B�ؖ���-----
certutil -addstore ROOT OutlookRecipientConfirmationAddin.cer

rem -----Visual Studio 2010 Tools for Office Runtime-----
vstor_redist.exe /Setup /passive /promptrestart

rem -----�C���X�g�[���[�Ăяo��-----
call addinInstaller.exe

rem -----dll���Ăяo��powershell�����s-----
powershell -NoProfile -ExecutionPolicy Unrestricted .\callRegistryDll.ps1