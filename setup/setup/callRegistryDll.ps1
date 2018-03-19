$ScriptPath = $(Split-Path -Parent $MyInvocation.MyCommand.Definition)
[Reflection.Assembly]::LoadFrom( $(Join-Path $ScriptPath "DoNotDisableAddinUpdater.dll.deploy"))
[DoNotDisableAddinUpdater.DoNotDisableAddinListUpdater]::UpdateDoNotDisableAddinList("OutlookRecipientConfirmationAddin", $TRUE)
