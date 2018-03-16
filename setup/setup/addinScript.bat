@echo off
cd /d %~dp0

rem -----DLLとPowerShellののブロック解除-----
start /wait streams.exe -d "DoNotDisableAddinUpdater.dll.deploy"
start /wait streams.exe -d "callRegistryDll.ps1"

rem -----セキュリティ証明書-----
certutil -addstore ROOT OutlookRecipientConfirmationAddin.cer

rem -----Visual Studio 2010 Tools for Office Runtime-----
vstor_redist.exe /Setup /passive /promptrestart

rem -----インストーラー呼び出し-----
call addinInstaller.exe

rem -----dllを呼び出すpowershellを実行-----
powershell -NoProfile -ExecutionPolicy Unrestricted .\callRegistryDll.ps1