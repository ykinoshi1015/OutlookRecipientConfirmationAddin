$ScriptPath = $(Split-Path -Parent $MyInvocation.MyCommand.Definition)
[Reflection.Assembly]::LoadFrom( $(Join-Path $ScriptPath "DoNotDisableAddinUpdaterDll.dll.deploy"))
$objoutlook = new-object -comobject outlook.application
[DoNotDisableAddinUpdaterDll.DoNotDisableAddinUpdaterDllClass]::checkDisable($objoutlook.version)
