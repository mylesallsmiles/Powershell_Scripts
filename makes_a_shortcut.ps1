$wshShell = New-Object -ComObject "WScript.Shell"
$urlShortcut = $wshShell.CreateShortcut(
  (Join-Path $wshShell.SpecialFolders.Item("AllUsersDesktop") "HelpDesk.url")
)
$urlShortcut.TargetPath = "#Link"
$urlShortcut.Save()

