@echo off
cd /d %~dp0

rem -----セキュリティ証明書-----
certutil -addstore ROOT OutlookRecipientConfirmationAddin.cer

rem -----Microsoft Visual C++ 2010 再頒布可能パッケージ (x86)-----
vcredist_x86.exe /Setup /passive /promptrestart

call addinInstall.exe